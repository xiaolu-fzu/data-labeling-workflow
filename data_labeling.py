"""Data Labeling Workflow v3.1"""
import os, sys, json, re, datetime, traceback

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(WORKSPACE, "input")

RUN_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_DIR = os.path.join(WORKSPACE, "runs", RUN_ID)
RUN_OUTPUT_DIR = os.path.join(RUN_DIR, "output")
RUN_GEN_DIR = os.path.join(RUN_DIR, "generated")

API_URL = None
API_KEY = None
MODEL = None
LOG_FILE = None


def log(msg):
    print(msg)
    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except:
            pass


def load_api_config():
    global API_URL, API_KEY, MODEL
    api_path = os.path.join(INPUT_DIR, "api.txt")
    if os.path.exists(api_path):
        with open(api_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip()
                if key == "api_base":
                    API_URL = val.rstrip("/") + "/chat/completions"
                elif key == "api_key":
                    API_KEY = val
                elif key == "model":
                    MODEL = val
    if not API_URL or not API_KEY or not MODEL:
        log("X API not configured in input/api.txt")
        return False
    return True


def read_config_txt(filepath):
    data_files = []
    requirement_file = None
    dp = "data_file: "
    rp = "requirement_file: "
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith(dp):
                name = line[len(dp):].strip().strip('"')
                if name:
                    data_files.append(name)
            elif line.startswith(rp):
                name = line[len(rp):].strip().strip('"')
                if name:
                    requirement_file = name
    return data_files, requirement_file


def read_excel_info(filepath):
    import pandas as pd
    try:
        if filepath.lower().endswith(".csv"):
            df = pd.read_csv(filepath, nrows=5)
            try:
                total_rows = sum(1 for _ in open(filepath, "r", encoding="utf-8")) - 1
            except:
                total_rows = "unknown"
        else:
            df = pd.read_excel(filepath, nrows=5)
            try:
                full = pd.read_excel(filepath, usecols=[0])
                total_rows = len(full)
            except:
                total_rows = "unknown"
        return {
            "filename": os.path.basename(filepath),
            "columns": df.columns.tolist(),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "estimated_rows": total_rows,
            "preview": df.head(5).to_dict(orient="records"),
        }
    except Exception as e:
        return {"filename": os.path.basename(filepath), "columns": [], "dtypes": {}, "estimated_rows": 0, "preview": []}


def read_requirement_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".txt" or ext == ".md":
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read()
    elif ext == ".xlsx" or ext == ".xls":
        import pandas as pd
        xls = pd.ExcelFile(filepath)
        parts = ["Labeling rules (ALL content from all sheets):"]
        for s in xls.sheet_names:
            parts.append("\n===== " + str(s) + " =====")
            df = pd.read_excel(filepath, sheet_name=s, dtype=str)
            parts.append("Columns: " + str(list(df.columns)))
            parts.append("Total rows: " + str(len(df)))
            for i in range(len(df)):
                vals = []
                for c in df.columns:
                    v = df.iloc[i][c]
                    if pd.notna(v):
                        vals.append(str(c) + ": " + str(v))
                if vals:
                    parts.append(str(i) + ": " + " | ".join(vals))
        return "\\n".join(parts)
    else:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()


def find_input_files():
    for f in os.listdir(INPUT_DIR):
        if f == "task_config.txt":
            cfg = os.path.join(INPUT_DIR, f)
            break
    else:
        return None, None, "Error: task_config.txt not found"
    data_files, req_file = read_config_txt(cfg)
    missing = []
    resolved = []
    for name in data_files:
        fp = os.path.join(INPUT_DIR, name)
        if os.path.exists(fp):
            resolved.append(fp)
        else:
            missing.append(name)
    rqf = None
    if req_file:
        fp = os.path.join(INPUT_DIR, req_file)
        if os.path.exists(fp):
            rqf = fp
        else:
            missing.append(req_file)
    if missing:
        return None, None, "Files not found: " + ", ".join(missing)
    return resolved, rqf, None


def call_llm(messages, max_tokens=16384, temperature=0.3):
    import urllib.request, ssl
    payload = json.dumps({"model": MODEL, "messages": messages, "max_tokens": max_tokens, "temperature": temperature}).encode("utf-8")
    req = urllib.request.Request(API_URL, data=payload, headers={"Authorization": "Bearer " + API_KEY, "Content-Type": "application/json"})
    ctx = ssl._create_unverified_context()
    resp = urllib.request.urlopen(req, context=ctx, timeout=180)
    result = json.loads(resp.read().decode("utf-8"))
    return result["choices"][0]["message"]["content"]


def extract_code(response):
    code = response
    if "```python" in code:
        code = code.split("```python")[1]
        if "```" in code:
            code = code.split("```")[0]
    elif "```" in code:
        code = code.split("```")[1]
        if "```" in code:
            code = code.split("```")[0]
    return code.strip()


def condense_rules(req_content):
    log("  Condensing labeling rules...")
    
    # Store original lines for post-check recovery
    original_lines = req_content.split("\n")
    
    prompt = (
        "You are a data processing assistant. Reorganize these labeling rules "
        "into a clear structured format.\n\n"
        "## RULES\n"
        "1. COMPRESS descriptive text only (explanations, notes, comments).\n"
        "2. PRESERVE ALL MAPPING TABLES completely: every row of character names, "
        "category codes, value mappings, or any table-like data must be kept VERBATIM.\n"
        "3. KEEP every field definition, every possible value, every label code.\n"
        "4. KEEP all decision logic, priority rules, edge cases.\n\n"
        "## Output structure\n\n"
        "=== LABEL_DEFINITIONS ===\n"
        "For each label field:\n"
        "  FIELD: name\n"
        "  TYPE: single|multi|text\n"
        "  ALL VALUES: (every value code + meaning)\n"
        "  LOGIC: (decision rules, priorities, conditions)\n"
        "  REQUIRES: (column names needed)\n"
        "  NOTES: (edge cases)\n\n"
        "=== MAPPING_TABLES ===\n"
        "ALL mapping/table data, EVERY row preserved verbatim.\n\n"
        "=== VALUE_MEANINGS ===\n"
        "Code -> full meaning + example\n\n"
        "WARNING: If you omit any mapping row, the labeling will be WRONG. "
        "Copy ALL rows from ALL tables.\n\n"
        "Rules:\n\n" + req_content
    )
    
    result = call_llm([{"role": "user", "content": prompt}], max_tokens=4096, temperature=0.1)
    
    # Universal post-check: find "data rows" in original that are missing from result
    # Data rows = lines with tab delimiters, or numbered items like "Row 1: ...", or lines containing "|"
    result_text = result
    result_set = set(result_text.split("\n"))
    
    missing_lines = []
    for orig_line in original_lines:
        stripped = orig_line.strip()
        if not stripped or len(stripped) < 10:
            continue
        # Look for data-row patterns
        is_data_row = (
            "\t" in stripped  # tab-separated data
            or stripped.count(" | ") >= 2  # pipe-separated
            or stripped.startswith("Row ")  # "Row N:" format
            or (":" in stripped[:5] and stripped[0].isdigit())  # "N: value" format
        )
        if not is_data_row:
            continue
        # Check if this line exists in result (fuzzy match)
        found = False
        for rl in result_set:
            if len(rl) < 10:
                continue
            # Use trigram overlap for fuzzy matching
            a_tri = set(stripped[i:i+3] for i in range(len(stripped)-2))
            b_tri = set(rl[i:i+3] for i in range(len(rl)-2))
            if len(a_tri) > 0 and len(a_tri & b_tri) / len(a_tri) > 0.3:
                found = True
                break
        if not found:
            missing_lines.append(orig_line)
    
    if missing_lines:
        log("  Recovered %d data rows lost during condensation" % len(missing_lines))
        result += "\n\n# === AUTO-RECOVERED (from original) ===\n"
        for ml in missing_lines:
            result += ml + "\n"
    
    log("  Rules: %d -> %d chars" % (len(req_content), len(result)))
    with open(os.path.join(RUN_GEN_DIR, "condensed_rules.txt"), "w", encoding="utf-8") as f:
        f.write(result)
    return result
def local_fix_code(code, err):
    fixed = code
    fixes = 0
    if "unterminated string literal" in err:
        m = re.search(r'line (\d+)', err)
        if m:
            arr = fixed.split("\n")
            ln = int(m.group(1))
            if ln - 1 < len(arr):
                bad = arr[ln - 1]
                if bad.count("'") % 2 == 1:
                    arr[ln - 1] = bad + "'"
                    fixes += 1
                elif bad.count('"') % 2 == 1:
                    arr[ln - 1] = bad + '"'
                    fixes += 1
                if fixes:
                    fixed = "\n".join(arr)
    if fixes == 0 and "was never closed" in err:
        m = re.search(r'line (\d+)', err)
        if m:
            arr = fixed.split("\n")
            ln = int(m.group(1))
            if ln - 1 < len(arr):
                bad = arr[ln - 1]
                for ch, cl in [("(", ")"), ("[", "]"), ("{", "}")]:
                    if bad.count(ch) > bad.count(cl):
                        arr[ln - 1] = bad + cl * (bad.count(ch) - bad.count(cl))
                        fixes += 1
                        fixed = "\n".join(arr)
                        break
    return fixed, fixes


def build_code_generation_prompt(tables_info, req_content):
    ind = INPUT_DIR.replace("\\", "/")
    od = RUN_OUTPUT_DIR.replace("\\", "/")
    lines = ["Write a Python script for data labeling.", "## Input Data"]

    is_multi = len(tables_info) > 1
    first_file = tables_info[0]["filename"]
    main_info = tables_info[0]

    if is_multi:
        lines.append("Multiple input files to MERGE.")
        for i, t in enumerate(tables_info):
            lines.append("  Table %d: %s" % (i+1, t["filename"]))
            lines.append("    Columns: " + str(t["columns"]))
            lines.append("    Rows: " + str(t["estimated_rows"]))
        join_keys = set()
        for t in tables_info:
            if "甯栧瓙閾炬帴" in t["columns"]:
                join_keys.add("甯栧瓙閾炬帴")
        if join_keys:
            lines.append("CRITICAL: Merge tables by " + ", ".join(sorted(join_keys)))
            lines.append("Table 1 is MAIN. Others LEFT JOIN.")
            lines.append("For labeling fields needing comment data, use merged comment info.")
    else:
        lines.append("File: " + first_file)

    lines.append("")
    lines.append("Input dir: INPUT_DIR (variable provided at runtime)")
    lines.append("Output dir: OUTPUT_DIR (variable provided at runtime)")
    lines.append("CRITICAL: Do NOT hardcode directory paths. Use the variables INPUT_DIR and OUTPUT_DIR directly.")
    lines.append("")
    lines.append("AVAILABLE columns: " + str(main_info["columns"]))
    lines.append("Sample (first 5 rows): " + str(main_info["preview"]))
    lines.append("")
    lines.append("IMPORTANT: Only the AVAILABLE columns listed above exist in the data file.")
    lines.append("If a rule requires a column NOT in AVAILABLE columns, SKIP that rule.")
    lines.append("Do NOT reference non-existent columns or files.")
    lines.append("")
    lines.append("## Labeling Rules")
    lines.append(req_content)
    lines.append("")
    lines.append("## Requirements")
    lines.append("- Read: pd.read_excel(os.path.join(input_dir, filename))")
    if is_multi:
        lines.append("- Read ALL tables, MERGE by join key, then apply labels")
    lines.append("- Apply rules, add new columns with Chinese names")
    lines.append("- Multi-select: join values with semicolon")
    lines.append("- Save at the END: df_posts.to_excel(os.path.join(OUTPUT_DIR, output_name), index=False)")
    lines.append("- Output name: [Labeled] " + first_file)
    lines.append("- Keep all original columns")
    lines.append("- End with: df_posts.to_excel(...) after ALL labeling is done")
    lines.append("- Use absolute paths")
    lines.append("")
    lines.append("## Code Quality")
    lines.append("- Verify all () [] {} are properly closed")
    lines.append("- Before output: compile(code, exec) to verify syntax")
    return "\n".join(lines)


def execute_with_retry(code, code_path, prompt, data_files, max_retries=3):
    import time
    try:
        for f in os.listdir(RUN_OUTPUT_DIR):
            try: os.remove(os.path.join(RUN_OUTPUT_DIR, f))
            except: pass
        time.sleep(0.3)
    except:
        pass

    attempt = 1
    current_code = code

    try:
        compile(current_code, "<gen>", "exec")
    except SyntaxError as _se:
        current_code, _ = local_fix_code(current_code, str(_se))
    except:
        pass

    while attempt <= max_retries:
        try:
            exec_globals = {"__file__": code_path, "__name__": "__main__",
                           "INPUT_FILES": data_files, "INPUT_DIR": INPUT_DIR, "OUTPUT_DIR": RUN_OUTPUT_DIR}
            exec(current_code, exec_globals)
            log("  OK - Complete! (attempt %d)" % attempt)
            if attempt > 1:
                with open(code_path, "w", encoding="utf-8") as f:
                    f.write(current_code)
            return True
        except (SyntaxError, IndentationError) as e:
            last_error = str(e)
            log("  X Attempt %d syntax: %s" % (attempt, str(last_error)[:200]))
            if attempt < max_retries:
                fixed, count = local_fix_code(current_code, last_error)
                if count > 0:
                    current_code = fixed
                    log("  Local fix: %d issue(s)" % count)
                    with open(os.path.join(RUN_GEN_DIR, "fix_v%d.py" % attempt), "w", encoding="utf-8") as f:
                        f.write(current_code)
                    attempt += 1
                    continue
                log("  -> LLM fix...")
                try:
                    fix_prompt = ("Fix this Python syntax error: " + last_error +
                        "\nCheck: all parens/brackets/braces closed, strings terminated, colons present. Output COMPLETE fixed script only.")
                    result = call_llm([
                        {"role": "user", "content": prompt},
                        {"role": "assistant", "content": current_code},
                        {"role": "user", "content": fix_prompt},
                    ])
                    current_code = extract_code(result)
                    with open(os.path.join(RUN_GEN_DIR, "fix_v%d_llm.py" % attempt), "w", encoding="utf-8") as f:
                        f.write(current_code)
                    log("  LLM fix generated")
                except Exception as fe:
                    log("  LLM fix failed: " + str(fe)[:80])
            attempt += 1
        except Exception as e:
            last_error = str(e)
            log("  X Attempt %d runtime: %s" % (attempt, str(last_error)[:200]))
            if attempt < max_retries:
                log("  -> LLM fix...")
                try:
                    fix_prompt = ("Fix this runtime error: " + last_error +
                        "\nCheck file permissions, data types, column names. Output COMPLETE fixed script only.")
                    result = call_llm([
                        {"role": "user", "content": prompt},
                        {"role": "assistant", "content": current_code},
                        {"role": "user", "content": fix_prompt},
                    ])
                    current_code = extract_code(result)
                    with open(os.path.join(RUN_GEN_DIR, "fix_v%d_llm.py" % attempt), "w", encoding="utf-8") as f:
                        f.write(current_code)
                    log("  LLM fix generated")
                except Exception as fe:
                    log("  LLM fix failed: " + str(fe)[:80])
            attempt += 1

    log("  X Failed after %d attempts" % max_retries)
    return False


def main():
    global LOG_FILE
    os.makedirs(RUN_OUTPUT_DIR, exist_ok=True)
    os.makedirs(RUN_GEN_DIR, exist_ok=True)

    LOG_FILE = os.path.join(RUN_DIR, "run.log")

    log("=" * 50)
    log("  Data Labeling Workflow v3.1")
    log("  Run: " + RUN_ID)
    log("  Dir: " + RUN_DIR)
    log("=" * 50)

    if not load_api_config():
        return

    data_files, req_file, error = find_input_files()
    if error:
        log("\nX " + error)
        return

    log("\nData files (%d):" % len(data_files))
    for f in data_files:
        log("  - " + os.path.basename(f))
    log("Requirement: " + os.path.basename(req_file))

    # Phase 1
    log("\n" + "-" * 50)
    log("Phase 1: Reading data")
    log("-" * 50)

    tables_info = []
    for f in data_files:
        info = read_excel_info(f)
        tables_info.append(info)
        log("  %s: %s" % (info["filename"], str(info["columns"])))

    req_content = read_requirement_file(req_file)
    log("  Rules: %d chars" % len(req_content))

    with open(os.path.join(RUN_GEN_DIR, "data_preview.json"), "w", encoding="utf-8") as f:
        json.dump({"tables": tables_info}, f, ensure_ascii=False, default=str)

    # Phase 2
    log("\n" + "-" * 50)
    log("Phase 2: Condensing rules")
    log("-" * 50)

    if len(req_content) > 2000:
        req_content = condense_rules(req_content)
    else:
        log("  Skipped (under 2000 chars)")

    # Phase 3
    log("\n" + "-" * 50)
    log("Phase 3: Generating code")
    log("-" * 50)

    prompt = build_code_generation_prompt(tables_info, req_content)
    log("  Prompt: %d chars" % len(prompt))

    with open(os.path.join(RUN_GEN_DIR, "last_prompt.txt"), "w", encoding="utf-8") as f:
        f.write(prompt)

    log("  Waiting for DeepSeek...")
    response = call_llm([{"role": "user", "content": prompt}], max_tokens=16384)
    generated_code = extract_code(response)
    log("  Code: %d chars" % len(generated_code))

    code_path = os.path.join(RUN_GEN_DIR, "generated_labeling.py")
    with open(code_path, "w", encoding="utf-8") as f:
        f.write(generated_code)

    # Phase 4
    log("\n" + "-" * 50)
    log("Phase 4: Executing")
    log("-" * 50)

    execute_with_retry(generated_code, code_path, prompt, data_files)

    # Output
    try:
        output_files = [f for f in os.listdir(RUN_OUTPUT_DIR) if not f.startswith("~$")]
    except:
        output_files = []

    if output_files:
        log("\nOutput:")
        for f in output_files:
            fp = os.path.join(RUN_OUTPUT_DIR, f)
            size = os.path.getsize(fp) / 1024
            log("  OK %s (%.1f KB)" % (f, size))

    log("\n" + "=" * 50)
    log("  Workflow complete")
    log("  Run dir: " + RUN_DIR)
    log("=" * 50)


if __name__ == "__main__":
    main()

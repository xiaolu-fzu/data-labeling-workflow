import os
import pandas as pd
import re
from datetime import datetime
from collections import Counter

def load_data(input_dir):
    """Load and merge input data files."""
    # Load posts data
    posts_file = os.path.join(input_dir, '非纯文本帖子（浏览量1000以上+剔除官号帖子）.xlsx')
    df_posts = pd.read_excel(posts_file)
    
    # Load comments data
    comments_file = os.path.join(input_dir, '非纯文本帖子-帖子下评论数据.xlsx')
    df_comments = pd.read_excel(comments_file)
    
    # Group comments by post link and aggregate
    comments_agg = df_comments.groupby('帖子链接').agg({
        '评论内容': lambda x: ' | '.join(x.dropna().astype(str)),
        '点赞数': lambda x: list(x.dropna().astype(int)),
        '评论者': lambda x: list(x.dropna().astype(str))
    }).reset_index()
    
    # Rename aggregated columns
    comments_agg.columns = ['帖子链接', '评论内容_聚合', '点赞数_列表', '评论者_列表']
    
    # Merge comments into posts
    df_posts = df_posts.merge(comments_agg, on='帖子链接', how='left')
    
    # Add comment count from original comments data
    comment_counts = df_comments.groupby('帖子链接').size().reset_index(name='评论数_实际')
    df_posts = df_posts.merge(comment_counts, on='帖子链接', how='left')
    df_posts['评论数_实际'] = df_posts['评论数_实际'].fillna(0).astype(int)
    
    return df_posts, df_comments

def build_character_map():
    """Build character name mapping from the provided table."""
    character_map = {
        'アンビー': '安比', 'アンビー・デマラ': '安比', '零号アンビー': '安比', 'ゼロ号アンビー': '安比',
        'Anby': '安比', 'Anby Demara': '安比', 'Soldier 0 Anby': '安比', 'Zero Anby': '安比',
        'ニコ': '妮可', 'ニコ・デマラ': '妮可', 'Nicole': '妮可', 'Nicole Demara': '妮可',
        'ビリー': '比利', 'ビリー・キッド': '比利', 'Billy': '比利', 'Billy Kid': '比利',
        'Billy Kidd': '比利', 'Starshine Billy': '比利',
        '猫又': '猫又', 'ネコマタ': '猫又', '猫宮又奈': '猫又', 'Nekomata': '猫又', 'Nekomiya Mana': '猫又',
        'クレタ': '珂蕾妲', 'クレタ・ベロボーグ': '珂蕾妲', 'Koleda': '珂蕾妲', 'Koleda Belobog': '珂蕾妲',
        'ベン': '本', 'ベン・ビガー': '本', 'Ben': '本', 'Ben Bigger': '本',
        'アンドー': '安东', 'アンドー・イワノフ': '安东', 'Anton': '安东', 'Anton Ivanov': '安东',
        'グレース': '格莉丝', 'グレース・ハワード': '格莉丝', 'Grace': '格莉丝', 'Grace Howard': '格莉丝',
        'エレン': '艾莲', 'エレン・ジョー': '艾莲', 'Ellen': '艾莲', 'Ellen Joe': '艾莲',
        'ライカン': '莱卡恩', 'フォン・ライカン': '莱卡恩', 'Lycaon': '莱卡恩', 'Von Lycaon': '莱卡恩',
        'カリン': '可琳', 'カリン・ウィクス': '可琳', 'Corin': '可琳', 'Corin Wickes': '可琳',
        'リナ': '丽娜', 'アレクサンドリナ': '丽娜', 'アレクサンドリナ・セバスチャン': '丽娜',
        'Rina': '丽娜', 'Alexandrina': '丽娜', 'Alexandrina Sebastiane': '丽娜',
        '蒼角': '苍角', 'そうかく': '苍角', 'Soukaku': '苍角',
        '11号': '11号', 'イレブン': '11号', '11': '11号', 'Soldier 11': '11号', 'S11': '11号',
        '朱鳶': '朱鸢', 'シュエン': '朱鸢', 'Zhu Yuan': '朱鸢',
        '青衣': '青衣', 'チンイー': '青衣', 'Qingyi': '青衣',
        'ジェーン': '简', 'ジェーン・ドゥ': '简', 'Jane': '简', 'Jane Doe': '简',
        'セス': '赛斯', 'セス・ローウェル': '赛斯', 'Seth': '赛斯', 'Seth Lowell': '赛斯',
        'シーザー': '凯撒', 'シーザー・キング': '凯撒', 'Caesar': '凯撒', 'Caesar King': '凯撒',
        'バーニス': '柏妮思', 'バーニス・ホワイト': '柏妮思', 'Burnice': '柏妮思', 'Burnice White': '柏妮思',
        'ルーシー': '露西', 'ルーシー・デ・モンテフィオ': '露西', 'Lucy': '露西', 'Lucy de Montefio': '露西',
        'パイパー': '派派', 'パイパー・ウィール': '派派', 'Piper': '派派', 'Piper Wheel': '派派',
        '雅': '星见雅', '星見雅': '星见雅', 'ホシミ・ミヤビ': '星见雅', 'Miyabi': '星见雅',
        'Hoshimi Miyabi': '星见雅', 'Miyabi Hoshimi': '星见雅',
        '悠真': '浅羽悠真', '浅羽悠真': '浅羽悠真', 'アサバ・ハルマサ': '浅羽悠真',
        'Harumasa': '浅羽悠真', 'Asaba Harumasa': '浅羽悠真',
        '柳': '月城柳', '月城柳': '月城柳', 'ツキシロ・ヤナギ': '月城柳', 'Yanagi': '月城柳',
        'Tsukishiro Yanagi': '月城柳',
        'ライト': '莱特', 'ライト・ヘルム': '莱特', 'Lighter': '莱特', 'Lighter Helm': '莱特',
        'ヒューゴ': '雨果', 'ヒューゴ・ヴラド': '雨果', 'Hugo': '雨果', 'Hugo Vlad': '雨果',
        'ビビアン': '薇薇安', 'ヴィヴィアン': '薇薇安', 'Vivian': '薇薇安',
        'イヴリン': '伊芙琳', 'イヴリン・シェヴァリエ': '伊芙琳', 'Evelyn': '伊芙琳',
        'Evelyn Chevalier': '伊芙琳',
        'アストラ': '耀嘉音', 'アストラ・ヤオ': '耀嘉音', 'Astra': '耀嘉音', 'Astra Yao': '耀嘉音',
        'トリガー': '扳机', '「トリガー」': '扳机', 'Trigger': '扳机',
        '儀玄': '仪玄', 'イーシェン': '仪玄', 'Yixuan': '仪玄',
        '橘福福': '橘福福', 'チー・フーフー': '橘福福', 'Ju Fufu': '橘福福', 'Fufu': '橘福福',
        '潘引壺': '潘引壶', 'パン・インフー': '潘引壶', 'Pan Yinhu': '潘引壶',
        'アリス': '爱丽丝', 'アリス・タイムフィールド': '爱丽丝', 'Alice': '爱丽丝',
        'Alice Timefield': '爱丽丝',
        '柚葉': '浮波柚叶', 'ユズハ': '浮波柚叶', 'Yuzuha': '浮波柚叶',
        'ポンペイ': '波可娜', 'ポコナ': '波可娜', 'Pulchra': '波可娜', 'Pulchra Fellini': '波可娜',
        '零号アンビー': '零号安比', 'ゼロ号アンビー': '零号安比', 'Soldier 0 Anby': '零号安比',
        'Zero Anby': '零号安比',
        'オフィス': '奥菲丝', '「ゴーストファイア」': '奥菲丝', 'Orphie': '奥菲丝', 'Ghostfire': '奥菲丝',
        'ルシア': '卢西娅', 'ルチア': '卢西娅', 'Lucia': '卢西娅',
        'イドハリ': '伊德海莉', 'アイスクイーン': '伊德海莉', 'Yidhari': '伊德海莉', 'Ice Queen': '伊德海莉',
        '琉音': '琉音', 'るいおん': '琉音', 'Rui On': '琉音', 'Rui': '琉音', 'Call Center': '琉音',
        '盤岳': '般岳', 'ばんがく': '般岳', 'Bangaku': '般岳', 'Fire Master': '般岳',
        '葉瞬光': '叶瞬光', 'よう しゅんこう': '叶瞬光', '瞬光': '叶瞬光', 'Ye Shunguang': '叶瞬光',
        'Shunguang': '叶瞬光', 'White Hair': '叶瞬光',
        'ショウ': '照', '照': '照', 'Shou': '照', 'Zhao': '照',
        '千夏': '千夏', 'すんな': '千夏', 'サンナ': '千夏', 'Sunna': '千夏', 'Sunny': '千夏',
        'アリア': '爱芮', 'アンジェル': '爱芮', 'Aria': '爱芮', 'Angel': '爱芮',
        '南宮羽': '南宫羽', 'なんきゅう はご': '南宫羽', '南宮': '南宫羽', 'Nangong Yu': '南宫羽',
        'Nangong': '南宫羽', 'Dancing': '南宫羽',
        'シシフ': '希希芙', 'ディーバ': '希希芙', 'Cissia': '希希芙', 'Diva': '希希芙',
        'プロミア': '普罗米娅', 'アイスエグゼキューター': '普罗米娅', 'Promeia': '普罗米娅',
        'Ice Executor': '普罗米娅',
        'スターライト・ビリー': '星徽・比利', 'スタービリー': '星徽・比利', 'ビリー SP': '星徽・比利',
        'Starlight Billy': '星徽・比利', 'Star Billy': '星徽・比利', 'Billy SP': '星徽・比利'
    }
    return character_map

def label_is_pr(row):
    """Label whether post is PR/commercial."""
    content = str(row.get('帖子内容', ''))
    
    # Check for explicit PR markers
    pr_patterns = [
        '#PR', '#広告', '【PR】', '提供', '#ad', '#sponsored',
        'プロモーション', '宣伝', 'タイアップ'
    ]
    
    for pattern in pr_patterns:
        if pattern in content:
            return '01-明确商单'
    
    # Check for commercial content signals
    commercial_patterns = [
        '予約受付', '福袋', '販売', '発売', '新商品', '購入',
        'お知らせ', 'ニュース', 'リリース', 'キャンペーン', 'セール',
        '割引', '特典', '限定', '予約', '申込'
    ]
    
    commercial_count = 0
    for pattern in commercial_patterns:
        if pattern in content:
            commercial_count += 1
    
    if commercial_count >= 2:
        return '02-疑似商单'
    
    # Check if publisher looks like brand/store
    publisher_id = str(row.get('发布者ID链接', ''))
    brand_patterns = [
        'shop', 'store', '官方', '公式', 'news', 'info',
        'prtimes', 'valuepress', 'atpress'
    ]
    
    for pattern in brand_patterns:
        if pattern in publisher_id.lower():
            if commercial_count >= 1:
                return '02-疑似商单'
    
    # Check for non-commercial signals
    non_commercial_patterns = [
        '描いた', 'イラスト', 'ファンアート', 'fanart', 'cosplay',
        'コスプレ', '手作り', 'ハンドメイド', '攻略', '実況'
    ]
    
    for pattern in non_commercial_patterns:
        if pattern in content.lower():
            return '03-非商单'
    
    # Default
    if len(content) < 10:
        return '99-无法判断'
    
    return '03-非商单'

def label_is_ai(row):
    """Label whether content is AI-generated."""
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    # Check for explicit AI labels
    ai_labels = [
        '#AIイラスト', '#AIart', '#AI生成', 'Made with AI', '#pixai',
        '#midjourney', '#novelai', '#stablediffusion', '#aiart',
        '#aigenerated', '#aiimage', '#texttoimage', '#dalle',
        'AIイラスト', 'AI生成', 'AI art', 'AI-generated'
    ]
    
    for pattern in ai_labels:
        if pattern.lower() in combined.lower():
            return '01-明确AI生成'
    
    # Check for AI suspicion signals in comments
    ai_suspicion_patterns = [
        'AI?', 'これAI', 'AIだ', 'AIっぽい', 'AIくさい',
        'AIこわい', 'AI嫌い', 'AI生成?'
    ]
    
    for pattern in ai_suspicion_patterns:
        if pattern in comments:
            return '02-疑似AI生成'
    
    # Check for human creation evidence
    human_patterns = [
        '描いた', '手描き', '手書き', '制作過程', '工程',
        'sketch', 'process', 'wip', 'work in progress',
        'コスプレ', 'cosplay', '実写', 'photo'
    ]
    
    for pattern in human_patterns:
        if pattern.lower() in content.lower():
            return '03-非AI生成'
    
    # Default: innocent until proven guilty
    return '03-非AI生成'

def label_image_category(row):
    """Label image creation category."""
    post_type = row.get('帖子类型（标签说明：1=仅视频；2=仅图片；3=图片+视频；）', 0)
    
    # Check if applicable
    if post_type == 1:  # Video only
        return '97-不适用'
    
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    # Priority order based on rules
    # 01-插画/绘画
    illustration_patterns = [
        'イラスト', '#art', '#fanart', '描いた', '塗り', '#illustration',
        'sketch', '落書き', 'デジタル', '#artwork', '#drawing',
        'ファンアート', 'fanart'
    ]
    
    # 02-Cosplay
    cosplay_patterns = [
        '#cosplay', '#コスプレ', 'cos', 'コス', '#zzzcosplay',
        'レイヤー', 'コスプレイヤー', 'model', 'コスプレ'
    ]
    
    # 03-Meme/梗图
    meme_patterns = [
        '#meme', '#ネタ', '#パロディ', '#ジョーク', 'meme',
        'ネタ', 'パロディ', 'ジョーク', 'wwww', '草'
    ]
    
    # 04-手工/模型
    craft_patterns = [
        '作った', '制作', '#フィギュア', '#ガレージキット', '#手作り',
        '#ぬい', '#アクスタ', '#ハンドメイド', '#フェルト', '手作り',
        'ハンドメイド', 'フェルト', 'ぬいぐるみ'
    ]
    
    # 05-摄影/风景
    photo_patterns = [
        '#photography', '#風景', '#景色', 'スクリーンショット',
        '#screenshot', '#capture', 'フォト', 'photo mode'
    ]
    
    # 06-漫画/短漫
    manga_patterns = [
        '漫画', 'マンガ', '#comic', '#4コマ', '4コマ',
        '短編', 'ストーリー', '#story'
    ]
    
    # Check patterns with priority
    for pattern in illustration_patterns:
        if pattern.lower() in combined.lower():
            return '01-插画/绘画'
    
    for pattern in cosplay_patterns:
        if pattern.lower() in combined.lower():
            return '02-Cosplay'
    
    for pattern in meme_patterns:
        if pattern.lower() in combined.lower():
            return '03-Meme/梗图'
    
    for pattern in craft_patterns:
        if pattern.lower() in combined.lower():
            return '04-手工/模型'
    
    for pattern in photo_patterns:
        if pattern.lower() in combined.lower():
            return '05-摄影/风景'
    
    for pattern in manga_patterns:
        if pattern.lower() in combined.lower():
            return '06-漫画/短漫（非meme）'
    
    # Check if it's a simple character name format (common for illustration posts)
    char_name_pattern = r'^[A-Za-z\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff\s/\・\-\.|]+\s*[#＃]'
    try:
        if re.match(char_name_pattern, content.strip()):
            return '01-插画/绘画'
    except re.error:
        pass
    
    # If no clear category but has content
    if len(content) > 20:
        return '90-其他图片二创'
    
    return '99-无法判断'

def label_video_category(row):
    """Label video creation category."""
    post_type = row.get('帖子类型（标签说明：1=仅视频；2=仅图片；3=图片+视频；）', 0)
    
    # Check if applicable
    if post_type == 2:  # Image only
        return '97-不适用'
    
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    # 01-攻略/教学
    guide_patterns = [
        '#攻略', '#ガイド', '#ビルド', '#編成', '#解説',
        '#おすすめ', '#立ち回り', '攻略', 'ガイド', '解説',
        '編成', '立ち回り', 'ビルド', 'おすすめ'
    ]
    
    # 02-实况/直播
    live_patterns = [
        '#実況', '#プレイ', '#ライブ', '#配信', '#生放送',
        '#Vtuber', '実況', '配信', 'ライブ', '生放送',
        'プレイ', 'ゲーム実況'
    ]
    
    # 03-混剪/MAD/AMV
    mad_patterns = [
        '#MAD', '#AMV', '#混ぜ', '#音MAD', 'MAD', 'AMV',
        '音MAD', '混剪', '編集', '#編集'
    ]
    
    # 04-手书/动画
    animation_patterns = [
        '#アニメ', '#手書', '#animation', '#作画', '#原画',
        '手書', '手描きアニメ', 'アニメ', 'animation',
        '作画', '原画'
    ]
    
    # 05-搞笑/短剧
    comedy_patterns = [
        '#コント', '#ネタ', '#パロディ', 'コント', '短劇',
        'ショート', '#ショート', 'コメディ', '#コメディ'
    ]
    
    # 06-壁纸/展示/静置型视频
    wallpaper_patterns = [
        '壁紙', '#壁紙', '展示', '#展示', 'デモ', '#demo',
        'walkthrough', 'ウォークスルー', '背景', '#background'
    ]
    
    # Check patterns with priority
    for pattern in animation_patterns:
        if pattern.lower() in combined.lower():
            return '04-手书/动画'
    
    for pattern in mad_patterns:
        if pattern.lower() in combined.lower():
            return '03-混剪/MAD/AMV'
    
    for pattern in guide_patterns:
        if pattern.lower() in combined.lower():
            return '01-攻略/教学'
    
    for pattern in live_patterns:
        if pattern.lower() in combined.lower():
            return '02-实况/直播'
    
    for pattern in comedy_patterns:
        if pattern.lower() in combined.lower():
            return '05-搞笑/短剧'
    
    for pattern in wallpaper_patterns:
        if pattern.lower() in combined.lower():
            return '06-壁纸/展示/静置型视频'
    
    # If no clear category but has content
    if len(content) > 20:
        return '90-其他视频二创'
    
    return '99-无法判断'

def label_content_theme(row):
    """Label content theme - multi-select."""
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    themes = []
    
    # Check for character theme
    character_patterns = [
        'キャラ', 'キャラクター', '推し', '好き', 'かわいい', 'かっこいい'
    ]
    
    for pattern in character_patterns:
        if pattern in combined:
            themes.append('01-角色主题')
            break
    
    # Check for scene/story theme
    scene_patterns = [
        'ストーリー', 'シナリオ', '物語', 'エピソード', '章',
        'マップ', 'ステージ', 'エリア', '地域', 'フィールド',
        'バージョン', 'Ver.', 'アップデート', '更新'
    ]
    
    for pattern in scene_patterns:
        if pattern in combined:
            themes.append('02-场景/剧情主题')
            break
    
    # Check for battle/gameplay theme
    battle_patterns = [
        '戦闘', 'バトル', 'コンボ', 'ダメージ', 'スキル',
        '装備', '武器', '編成', 'パーティ', 'チーム',
        '攻略', 'クリア', '挑戦', 'スコア', 'ランキング'
    ]
    
    for pattern in battle_patterns:
        if pattern in combined:
            themes.append('03-战斗/玩法主题')
            break
    
    # If no theme detected
    if not themes:
        if len(combined) > 10:
            themes.append('90-其他主题')
        else:
            themes.append('99-无法判断')
    
    return ';'.join(themes)

def label_character_count(row):
    """Label number of characters mentioned."""
    content_theme = row.get('内容主题-角色/场景', '')
    
    # If theme is not character-focused, mark as not applicable
    if '01-角色主题' not in content_theme:
        return '97-不适用'
    
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    character_map = build_character_map()
    
    # Find all mentioned characters
    mentioned_chars = set()
    for keyword, chinese_name in character_map.items():
        if keyword in combined:
            mentioned_chars.add(chinese_name)
    
    # Also check for character names directly
    for chinese_name in set(character_map.values()):
        if chinese_name in combined:
            mentioned_chars.add(chinese_name)
    
    # Check for CP patterns
    cp_patterns = ['×', 'x', '&', 'と', 'カップル', 'CP']
    
    for pattern in cp_patterns:
        if pattern in combined:
            return '02-角色×角色（CP）'
    
    # Determine count
    if len(mentioned_chars) == 0:
        # Try to find character names not in the map
        name_pattern = r'[A-Z][a-z]+(?:\s[A-Z][a-z]+)*'
        try:
            names = re.findall(name_pattern, combined)
            if len(names) == 0:
                return '99-无法判断'
            elif len(names) == 1:
                return '01-单一角色'
            elif len(names) == 2:
                return '02-角色×角色（CP）'
            else:
                return '03-多角色'
        except re.error:
            return '99-无法判断'
    elif len(mentioned_chars) == 1:
        return '01-单一角色'
    elif len(mentioned_chars) == 2:
        return '02-角色×角色（CP）'
    else:
        return '03-多角色'

def label_mentioned_characters(row):
    """Extract mentioned character names."""
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    combined = content + ' ' + comments
    
    character_map = build_character_map()
    
    # Find all mentioned characters with their positions
    mentions = []
    for keyword, chinese_name in character_map.items():
        if keyword in combined:
            idx = combined.index(keyword)
            mentions.append((idx, chinese_name))
    
    # Sort by position in text
    mentions.sort(key=lambda x: x[0])
    
    # Get unique characters in order of appearance
    seen = set()
    ordered_chars = []
    for _, name in mentions:
        if name not in seen:
            seen.add(name)
            ordered_chars.append(name)
    
    if not ordered_chars:
        return ''
    
    # Format based on count
    if len(ordered_chars) == 1:
        return ordered_chars[0]
    elif len(ordered_chars) == 2:
        # Sort by Japanese 五十音 order (using Unicode as approximation)
        sorted_chars = sorted(ordered_chars)
        return f'{sorted_chars[0]}×{sorted_chars[1]}'
    else:
        # Return the most core character (first mentioned)
        return ordered_chars[0]

def label_comment_motivation(row, df_comments):
    """Label main motivation of comments."""
    post_link = row.get('帖子链接', '')
    
    # Get comments for this post
    post_comments = df_comments[df_comments['帖子链接'] == post_link]
    
    if len(post_comments) <= 1:
        return '99-无法判断'
    
    # Get top comments by likes (≥3 likes, top 5)
    top_comments = post_comments[post_comments['点赞数'] >= 3].nlargest(5, '点赞数')
    
    if len(top_comments) == 0:
        return '99-无法判断'
    
    motivations = []
    
    for _, comment_row in top_comments.iterrows():
        comment_text = str(comment_row.get('评论内容', ''))
        
        # Check for criticism/controversy (highest priority)
        criticism_patterns = [
            'きもい', '不快', '通報', '気持ち悪い', '消えろ',
            '削除', 'クレーム', '低質', 'つまらない', 'AIこわい',
            'AI嫌い', 'ひどい', '最悪', '下手', 'へたくそ'
        ]
        
        found_criticism = False
        for pattern in criticism_patterns:
            if pattern in comment_text:
                motivations.append('05-批评/争议')
                found_criticism = True
                break
        
        if not found_criticism:
            # Check for praise
            praise_patterns = [
                'かわいい', 'すごい', '美しい', '最高', '好き',
                'カッコいい', '素敵', 'セクシー', '美し', '神',
                '❤️', '😍', '素晴らしい', '凄い', '可愛い'
            ]
            
            if any(p in comment_text for p in praise_patterns):
                motivations.append('01-赞美/夸奖')
            
            # Check for meme/interaction
            meme_patterns = ['wwww', '草', 'ww', '笑った', '草生える']
            
            if any(p in comment_text for p in meme_patterns):
                motivations.append('02-玩梗/互动')
            
            # Check for question/help
            question_patterns = [
                'どうやって', '教えて', 'どこ', 'いくら', 'これ何',
                '教えてください', '質問', '教えて'
            ]
            
            if any(p in comment_text for p in question_patterns):
                motivations.append('03-询问/求助')
            
            # Check for thanks/support
            thanks_patterns = [
                'ありがとう', '感謝', 'お疲れ', '応援', 'ファイト',
                '♥', 'フォローしました', 'すてき'
            ]
            
            if any(p in comment_text for p in thanks_patterns):
                motivations.append('04-感谢/应援')
            
            # Check for info supplement
            info_patterns = [
                'ここ違います', '正しくは', '補足', '追加情報',
                'この情報は古い', '違う', '間違い'
            ]
            
            if any(p in comment_text for p in info_patterns):
                motivations.append('07-情报补充/纠错')
            
            # Check for casual chat
            chat_patterns = [
                'ゲームフレンド募集', '一緒に', 'お話しましょう',
                'フレンド', '友達'
            ]
            
            if any(p in comment_text for p in chat_patterns):
                motivations.append('06-闲聊/社交')
            
            # Check for check-in
            checkin_patterns = [
                'おはよう', 'おやすみ', 'お疲れ様', '1コメ',
                '初めまして', 'こんにちは', 'こんばんは'
            ]
            
            if any(p in comment_text for p in checkin_patterns):
                motivations.append('08-打卡/标记')
    
    # If no motivation detected
    if not motivations:
        return '90-其他动机'
    
    # If criticism exists, prioritize it
    if '05-批评/争议' in motivations:
        return '05-批评/争议'
    
    # Return most frequent motivation
    motivation_counts = Counter(motivations)
    most_common = motivation_counts.most_common(1)[0][0]
    
    # If info supplement and praise both exist, return both
    if '07-情报补充/纠错' in motivations and '01-赞美/夸奖' in motivations:
        return '01-赞美/夸奖;07-情报补充/纠错'
    
    return most_common

def label_controversy(row, df_comments):
    """Label controversy signal."""
    post_link = row.get('帖子链接', '')
    
    # Get comments for this post
    post_comments = df_comments[df_comments['帖子链接'] == post_link]
    
    if len(post_comments) == 0:
        return '99-无法判断'
    
    # Check for negative signals
    negative_patterns = [
        'きもい', '不快', '通報', '気持ち悪い', '消えろ',
        '削除', 'クレーム', 'AIこわい', 'AI嫌い', 'ひどい',
        '最悪', '下手', 'へたくそ', '低質', 'つまらない'
    ]
    
    for _, comment_row in post_comments.iterrows():
        comment_text = str(comment_row.get('评论内容', ''))
        likes = comment_row.get('点赞数', 0)
        
        for pattern in negative_patterns:
            if pattern in comment_text and likes >= 3:
                return '01-有争议'
    
    return '02-无争议'

def label_information_sufficiency(row):
    """Label information sufficiency."""
    content = str(row.get('帖子内容', ''))
    comments = str(row.get('评论内容_聚合', ''))
    
    # Count how many fields would be "无法判断"
    undetermined_count = 0
    
    # Check if we have enough content
    if len(content) < 20:
        undetermined_count += 1
    
    # Check if we have comments
    if not comments or len(comments) < 10:
        undetermined_count += 1
    
    # Check for key information
    has_key_info = bool(re.search(r'[A-Za-z\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff]{2,}', content))
    if not has_key_info:
        undetermined_count += 1
    
    if undetermined_count <= 1:
        return '01-充分'
    else:
        return '02-不充分'

def main(input_dir, output_dir):
    """Main function to run the labeling pipeline."""
    print("Loading data...")
    df_posts, df_comments = load_data(input_dir)
    
    print(f"Loaded {len(df_posts)} posts and {len(df_comments)} comments")
    
    # Apply labeling rules
    print("Applying labels...")
    
    # 是否商单（PR）
    print("  - Labeling PR status...")
    df_posts['是否商单（PR）'] = df_posts.apply(label_is_pr, axis=1)
    
    # 是否AI生成
    print("  - Labeling AI generation...")
    df_posts['是否AI生成'] = df_posts.apply(label_is_ai, axis=1)
    
    # 二创类别（图片类）
    print("  - Labeling image category...")
    df_posts['二创类别（图片类）'] = df_posts.apply(label_image_category, axis=1)
    
    # 二创类别（视频类）
    print("  - Labeling video category...")
    df_posts['二创类别（视频类）'] = df_posts.apply(label_video_category, axis=1)
    
    # 内容主题-角色/场景
    print("  - Labeling content theme...")
    df_posts['内容主题-角色/场景'] = df_posts.apply(label_content_theme, axis=1)
    
    # 提及角色数量
    print("  - Labeling character count...")
    df_posts['提及角色数量'] = df_posts.apply(label_character_count, axis=1)
    
    # 提及角色
    print("  - Labeling mentioned characters...")
    df_posts['提及角色'] = df_posts.apply(label_mentioned_characters, axis=1)
    
    # 用户评论主流动机
    print("  - Labeling comment motivation...")
    df_posts['用户评论主流动机'] = df_posts.apply(
        lambda row: label_comment_motivation(row, df_comments), axis=1
    )
    
    # 争议信号
    print("  - Labeling controversy...")
    df_posts['争议信号'] = df_posts.apply(
        lambda row: label_controversy(row, df_comments), axis=1
    )
    
    # 信息充分度
    print("  - Labeling information sufficiency...")
    df_posts['信息充分度'] = df_posts.apply(label_information_sufficiency, axis=1)
    
    # Remove temporary columns
    columns_to_drop = ['评论内容_聚合', '点赞数_列表', '评论者_列表', '评论数_实际']
    df_posts = df_posts.drop(columns=[c for c in columns_to_drop if c in df_posts.columns])
    
    # Save output
    output_name = '[Labeled] 非纯文本帖子（浏览量1000以上+剔除官号帖子）.xlsx'
    output_path = os.path.join(output_dir, output_name)
    
    print(f"Saving to {output_path}...")
    df_posts.to_excel(output_path, index=False)
    
    print(f"Done! Labeled {len(df_posts)} posts.")
    print(f"Output saved to: {output_path}")
    
    return df_posts

# The script will be called with INPUT_DIR and OUTPUT_DIR variables
if __name__ == "__main__":
    # These variables will be provided at runtime
    # INPUT_DIR = "path/to/input"
    # OUTPUT_DIR = "path/to/output"
    
    # For testing, you can uncomment and set these:
    # INPUT_DIR = "./input"
    # OUTPUT_DIR = "./output"
    
    df_result = main(INPUT_DIR, OUTPUT_DIR)
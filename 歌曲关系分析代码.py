import pandas as pd
from collections import defaultdict
import json
from excel_api import Workbook

# 定义关系类型词汇词典
relationship_types = {
    '父母': ['爸爸', '父亲', '妈妈', '母亲', '爸妈', '双亲', '父母', '爹', '娘'],
    '配偶/伴侣': ['爱人', '恋人', '伴侣', '配偶', '妻子', '丈夫', '老婆', '老公', '对象', '另一半', '情侣'],
    '子女': ['儿子', '女儿', '孩子', '子女', '小孩', '宝贝', '宝宝'],
    '朋友': ['朋友', '友人', '兄弟', '姐妹', '哥们', '闺蜜', '兄弟伙', '小伙伴'],
    '自己': ['自己', '我', '我自己', '自身', '本人', '自我'],
    '爱情': ['爱情', '爱', '相恋', '相爱', '恋爱', '情深', '深情', '情投意合', '心意相通'],
    '亲情': ['亲情', '温情', '温暖', '家人', '家庭', '家', '阖家', '团圆', '团聚'],
    '友情': ['友情', '友谊', '交情', '挚交', '莫逆之交', '刎颈之交'],
    '陌生人': ['陌生人', '路人', '路过', '过客', '旁人', '他人', '别人', '路人甲']
}

# 读取Excel文件并进行关系提取分析
def analyze_relationships(sheet_name, sheet):
    # 创建计数器
    relationship_counts = defaultdict(int)
    
    # 获取所有数据行
    max_row = sheet.get_max_row_with_data_in_column("B")
    
    # 分批读取数据以避免内存问题
    batch_size = 100
    for batch_start in range(2, max_row + 1, batch_size):
        batch_end = min(batch_start + batch_size - 1, max_row)
        range_str = f"A{batch_start}:B{batch_end}"
        
        try:
            # 获取当前批次的数据
            data = sheet.get_raw_range_data(range_str)
            
            for row in data:
                if len(row) >= 2 and row[1]:  # 确保有分词结果
                    分词_result = str(row[1])
                    
                    # 检查每个关系类型
                    for relationship_type, keywords in relationship_types.items():
                        for keyword in keywords:
                            if keyword in 分词_result:
                                relationship_counts[relationship_type] += 1
        except Exception as e:
            print(f"处理批次 {batch_start}-{batch_end} 时出错: {e}")
    
    return relationship_counts

# 打开Excel文件
wb = Workbook("/app/data/files/男女歌手歌曲分词结果.xlsx")

# 分析男歌手歌曲
male_sheet = wb.get_sheet("男歌手歌曲")
male_relationships = analyze_relationships("男歌手歌曲", male_sheet)

# 分析女歌手歌曲
female_sheet = wb.get_sheet("女歌手歌曲")
female_relationships = analyze_relationships("女歌手歌曲", female_sheet)

# 关闭工作簿
wb.close()

# 输出结果
print("男歌手歌曲关系类型统计:")
for relationship_type, count in sorted(male_relationships.items(), key=lambda x: x[1], reverse=True):
    print(f"{relationship_type}: {count}")

print("\n女歌手歌曲关系类型统计:")
for relationship_type, count in sorted(female_relationships.items(), key=lambda x: x[1], reverse=True):
    print(f"{relationship_type}: {count}")

# 保存结果到文件，以便后续生成图表
results = {
    "male_relationships": dict(male_relationships),
    "female_relationships": dict(female_relationships),
    "relationship_types": relationship_types
}

with open("/app/data/files/男女歌手歌曲关系分析结果.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\n结果已保存到 男女歌手歌曲关系分析结果.json")
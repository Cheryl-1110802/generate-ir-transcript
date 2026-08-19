# 職位和姓名常數
CHAIRMAN = "徐清祥"
PRESIDENT = "何明洲"
FINANCIAL_OFFICER = "夏喬威"
CHAIRMAN_EN = "Dr. Charles Hsu"
PRESIDENT_EN = "Michael Ho"
FINANCIAL_OFFICER_EN = "Joseph Hsia"

# 名字提取功能
def get_first_name(full_name):
    """從全名中提取第一個名字，忽略稱謂"""
    if not full_name:
        return ""
    
    # 移除常見稱謂
    name = (full_name.replace("Dr. ", "")
                    .replace("Mr. ", "")
                    .replace("Ms. ", "")
                    .replace("Mrs. ", "")
                    .replace("Prof. ", "")
                    .strip())
    
    # 取第一個單字
    return name.split()[0] if name else ""

# 提取的英文名字常數
CHAIRMAN_FIRST_NAME = get_first_name(CHAIRMAN_EN)      # "Charles"
PRESIDENT_FIRST_NAME = get_first_name(PRESIDENT_EN)    # "Michael"
FINANCIAL_OFFICER_FIRST_NAME = get_first_name(FINANCIAL_OFFICER_EN)  # "Joseph"

import pandas as pd

# 生成测试数据
prizes = {
    '奖项': ['一等奖', '二等奖', '三等奖'],
    '数量': [1, 3, 5]
}

participants = {
    '姓名': [
        '张三', '李四', '王五', '赵六', '孙七',
        '周八', '吴九', '郑十', '钱十一', '孙十二',
        '李十三', '王十四', '赵十五', '孙十六', '周十七',
        '吴十八', '郑十九', '钱二十', '孙二十一', '李二十二'
    ]
}

# 创建DataFrame
prizes_df = pd.DataFrame(prizes)
participants_df = pd.DataFrame(participants)

# 保存到Excel
with pd.ExcelWriter('lottery_data.xlsx') as writer:
    prizes_df.to_excel(writer, sheet_name='prizes', index=False)
    participants_df.to_excel(writer, sheet_name='participants', index=False)

print("测试数据已生成并保存为 lottery_data.xlsx")

import openpyxl


def excel_to_markdown(excel_filepath):
    """
    使用openpyxl读取excel生成markdown，文件名为1级标题
    读取excel以下这5列，相关研发需求为二级标题（视为类别，多条记录需求相同将只有一个二级（下方依次显示三级用例）），
    用例标题为三级标题，前置条件为四级标题，步骤按换行分割成使用无序列表格式，预期使用步骤下一级无序列表格式
    :param excel_filepath: your_test_case.xlsx
    :return:
    """

    # 载入Excel文件
    wb = openpyxl.load_workbook(excel_filepath)
    sheet = wb.active

    # 创建一个空字符串用来保存Markdown内容
    markdown_content = ""

    # 获取文件名并作为一级标题
    file_name = excel_filepath[:excel_filepath.rfind(".")]
    markdown_content += f"# {file_name}\n\n"

    # 按需求或者模块进行划分
    related_demand_list = []

    # 读取表格内容
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # 如果某一行数据的列数大于5列，避免解包错误，使用 list(row) 来处理
        row_data = list(row)

        # 确保该行至少包含5列
        if len(row_data) >= 5:
            related_demand = row_data[0]
            case_title = row_data[1]
            pre_condition = row_data[2]
            steps = row_data[3]
            expected = row_data[4]

            # 相关研发需求作为二级标题
            if related_demand and related_demand not in related_demand_list:
                related_demand_list.append(related_demand)
                markdown_content += f"## {related_demand}\n"

            # 用例标题作为三级标题
            if case_title is None: continue  # 没有标题直接跳过
            markdown_content += f"### {case_title}\n"

            # 前置条件作为四级标题
            if pre_condition: markdown_content += f"#### {pre_condition}\n\n"

            # 步骤与预期
            if steps:
                steps_list = steps.split('\n')
                for step in steps_list:
                    markdown_content += f"- {step}\n"

            # 预期结果放在下一级无序列表中
            if expected:
                expected_list = expected.split('\n')
                for expect in expected_list:
                    markdown_content += f"  - {expect}\n"

            markdown_content += "\n"  # 添加空行分隔每个用例

    # 保存结果到Markdown文件
    with open(file_name + ".md", "w", encoding="utf-8") as md_file:
        md_file.write(markdown_content)


# 调用函数并传入Excel文件路径
excel_to_markdown('1markdown转成的excel.xlsx')

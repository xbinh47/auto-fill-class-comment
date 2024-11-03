import pandas as pd

list_label = [
    "Nội dung buổi học",
    "Link slide bài giảng",
    "Link Video hướng dẫn",
    "Yêu cầu cho buổi tiếp theo",
    "Hạn nộp bài",
    "Nội dung buổi học tới",
    "📌Tình hình học tập của lớp",
]

extra_label = {
    "class_performance": "📌Tình hình học tập của lớp",
    "deadline": "⏰Hạn nộp bài",
}

def read_lesson_and_format_from_excel(file_path, sheet_name, lesson_number):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Determine the column name based on the lesson number
    lesson_column = f"Buổi {lesson_number}"
    
    # Check if the column exists
    if lesson_column not in df.columns:
        raise ValueError(f"Lesson {lesson_number} not found in sheet {sheet_name}.")
    
    # Extract the data from the specified column
    data = df[lesson_column].dropna().tolist()
    lesson_datas = []
    for lesson_data in data:
        # check lesson data cotain list_label
        if any(label in lesson_data for label in list_label):
            if ":" in lesson_data and lesson_data.split(":")[1].strip() != "":
                parts = lesson_data.split(":", 1)
                if len(parts) == 2:
                    lesson_datas.append(parts[0].strip()) 
                    lesson_datas.append(parts[1].strip()) 
        else:
            lesson_datas.append(lesson_data.replace("\n", ""))
    
    return lesson_datas

data = read_lesson_and_format_from_excel("data.xlsx", "PTA", 1)
print(data)

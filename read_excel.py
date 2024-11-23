import pandas as pd

list_label = [
    "Nội dung buổi học",
    "Link slide bài giảng",
    "Link Video hướng dẫn",
    "Yêu cầu cho buổi tiếp theo",
    "Hạn nộp bài",
    "Nội dung buổi học tới",
    "📌Tình hình học tập của lớp",
]

extra_label = {
    "class_performance": "📌Tình hình học tập của lớp",
    "homework_result": "🏆Kết quả bài tập về nhà",
    "deadline": "⏰Hạn nộp bài",
    "next_requirement": "🔍Yêu cầu cho buổi tiếp theo",
}

def read_lesson_and_format_from_excel(file_path, sheet_name, lesson_number):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    
    # Determine the column name based on the lesson number
    lesson_column = f"Buổi {lesson_number}"
    
    # Check if the column exists
    if lesson_column not in df.columns:
        raise ValueError(f"Lesson {lesson_number} not found in sheet {sheet_name}.")
    
    # Extract the data from the specified column
    data = df[lesson_column].tolist()
    lesson_datas = []
    seen_data = set()  # Track seen data to avoid duplicates

    for lesson_data in data:
        # Convert lesson_data to string to handle NaN values
        lesson_data_str = str(lesson_data)

        # Check if any label is in the lesson_data_str
        if any(label in lesson_data_str for label in list_label):
            lesson_data_str = lesson_data_str.replace("\n", "")
            # Get the index of the current lesson_data
            current_index = df.index[df[lesson_column] == lesson_data].tolist()
            if current_index:
                # Safely access the next row
                next_index = current_index[0] + 1
                if next_index < len(df):
                    next_row = df.loc[next_index, lesson_column]
                    # Append if next_row is valid and not seen
                    if not pd.isna(next_row) and lesson_data_str not in seen_data:
                        lesson_datas.append(lesson_data_str)
                        seen_data.add(lesson_data_str)
        else:
            # Directly append if no label is found and not seen
            if lesson_data_str not in seen_data:
                lesson_datas.append(lesson_data_str)
                seen_data.add(lesson_data_str)

    return clean_nan(lesson_datas)
    
def clean_nan(data):
    return [item for item in data if item != "nan"]
    
data = read_lesson_and_format_from_excel("data.xlsx", "PTA", 9)
print(data)

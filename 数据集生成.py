import random
import csv
import os

random.seed(1)  # 设置随机数种子

def generate_data(num_samples, start=0):
    data = []
    for _ in range(start, num_samples):
        l4 = round(random.uniform(20, 39), 1)
        l3 = round(random.uniform(40, 59), 1)
        l2 = round(random.uniform(60, 79), 1)
        l1 = round(random.uniform(80, 100), 1)
        d1 = round(random.uniform(20, 40), 1)
        d2 = round(random.uniform(15, 35), 1)
        h = round(random.uniform(0.3, 0.7), 1)
        data.append((l4, l3, l2, l1, d1, d2, h))
    return data

def append_data_to_csv(file_name, data, header=None):
    if not file_name.endswith('.csv'):
        raise ValueError("文件名必须以'.csv'结尾")
    
    file_exists = os.path.isfile(file_name)
    with open(file_name, 'a', newline='') as file:
        writer = csv.writer(file)
        if not file_exists and header:
            writer.writerow(header)
        writer.writerows(data)

def count_data_samples(csv_file):
    try:
        with open(csv_file, 'r') as file:
            return sum(1 for row in csv.reader(file)) - 1  # 减去表头
    except FileNotFoundError:
        return 0

# 目标生成的总样本数
total_num_samples = 1000
csv_file_name = 'data_set.csv'
existing_samples = count_data_samples(csv_file_name)

# 计算还需要生成多少数据
samples_to_generate = total_num_samples - existing_samples

if samples_to_generate > 0:
    data_set = generate_data(total_num_samples, start=existing_samples)
    header = ['l4', 'l3', 'l2', 'l1', 'd1', 'd2', 'h'] if existing_samples == 0 else None
    append_data_to_csv(csv_file_name, data_set, header=header)
    print(f"新增{samples_to_generate}个样本，数据已追加保存为{csv_file_name}文件")
else:
    print("数据集已达到或超过目标样本数，无需生成更多数据。")

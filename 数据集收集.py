# -----------------------------------------------------------------------------
# 导入所需模块
# -----------------------------------------------------------------------------
import numpy as np
import pandas as pd
import cst
import time 
import random
import csv
import os
import glob
import re
from cst.interface import DesignEnvironment 
import matplotlib.pyplot as plt
import psutil
import sys
import pygetwindow as gw
from PIL import Image
import threading
import pygetwindow as gw
import win32gui

# -----------------------------------------------------------------------------
# 选择文件根目录
# -----------------------------------------------------------------------------
path = r"C:\Users\CHEESZ\Desktop\Test"
# 创建文件夹
folders = ["CST_DOC", "CST_OUTPUT", "PIC", "PROGRESS"]
folder_paths = []
for folder in folders:
    folder_path = os.path.join(path, folder)
    os.makedirs(folder_path, exist_ok=True)
    folder_paths.append(folder_path)
print("文件夹已创建：CST_DOC, CST_OUTPUT, PIC, PROGRESS")
print("文件夹路径：", folder_paths)
# -----------------------------------------------------------------------------
# CST库环境设置
# -----------------------------------------------------------------------------
def import_from_previous():
    if os.path.exists('previous_import.txt'):
        with open('previous_import.txt', 'r') as file:
            previous_import = file.read()

            # 使用 threading.Event 来实现超时
            user_input = threading.Event()
            answer = None  # 初始化 answer 变量

            def get_user_choice():
                nonlocal user_input, answer
                answer = input('是否使用上次导入的地址 ' + previous_import + ' ? (y/n): ')
                user_input.set()  # 用户输入完成，设置 Event 信号

            # 创建并启动线程
            user_input_thread = threading.Thread(target=get_user_choice)
            user_input_thread.start()
            time.sleep(0.5)
            print('\n')
            # 倒计时等待用户输入，最多等待 10 秒
            for i in range(2, -1, -1):
                if user_input.is_set():
                    break
                print(f"还有{i}s时间决定")
                time.sleep(1)

            # 判断用户是否输入完成
            if user_input.is_set():
                # 用户输入完成，返回相应的结果
                if answer.lower() == 'y':
                    return previous_import
                elif answer.lower() == 'n':
                    return None

            # 用户未在超时时间内输入，自动使用上次导入的地址
            return previous_import

    return None


def save_previous_import(address):
    '''保存最后一次导入的地址'''

    with open('previous_import.txt', 'w') as file:
        file.write(address)

def get_import_address():
    '''返回导入地址'''

    previous_import = import_from_previous()
    if previous_import:
        return previous_import

    address = input('请输入cst库地址(如E:\\CST Studio Suite 2020\\AMD64\\python_cst_libraries)：')
    save_previous_import(address)
    return address

# 执行
address = get_import_address()
if '%s'%address not in sys.path:
    sys.path.append('%s'%address)
sys.path = list(set(sys.path))
# -----------------------------------------------------------------------------
# CST常用函数
# -----------------------------------------------------------------------------

def brick(name,component,material,xrange,yrange,zrange):
    line_break = '\n'#换行
    xmin, xmax = xrange
    ymin, ymax = yrange
    zmin, zmax = zrange
    sCommand = ['With Brick',
            '.Reset',
            '.Name "%s" '%name,
            '.Component "%s" '%component,
            '.Material "%s"'%material ,
            '.Xrange "%s", "%s"' %(xmin,xmax),
            '.Yrange "%s", "%s"'%(ymin,ymax) ,
            '.Zrange "%s", "%s"'%(zmin,zmax) ,
            '.Create',
            'End With'] 
    sCommand = line_break.join(sCommand)
    modelers.add_to_history('Brick', sCommand)

def ChangeColour(name,R,G,B):
    sCommand = f'''With Material 
        .Name "{name}"
        .Folder ""
        .Colour "{R}", "{G}", "{B}" 
        .Wireframe "False" 
        .Reflection "False" 
        .Allowoutline "True" 
        .Transparentoutline "False" 
        .Transparency "0" 
        .ChangeColour 
        End With'''
    modelers.add_to_history('CC', sCommand)

def wcs_face(component:str,name:str,faceid:int) -> None:
    line_break = '\n'
    sCommand = ['Pick.PickFaceFromId("%s:%s", "%d" )'%(component,name,faceid),'WCS.AlignWCSWithSelected "Face"']
    sCommand = line_break.join(sCommand)
    modelers.add_to_history('wcs_face', sCommand)

def exdata(sp:str,type:str,format:str,path:str,name:str):
    '''
    This method is uesd to export sim data

    Args:
        sp(str):SZmax(1),Zmax(1)\n
        type(str):mag,dB,real,imag,phase
        format(str):txt,csv
    '''
    line_break = '\n'
    tree = "1D Results\\S-Parameters\\" + sp
    if type == "mag":
        sCommand=['SelectTreeItem ("%s")'%tree,
            'With Plot1D',
            '.PlotView "magnitude"',
            '.Plot',
            'End With',]
        sCommand = line_break.join(sCommand)
        modelers.add_to_history('mag', sCommand)
    else:
        print("请输入mag/dB/real/imag/phase")
    
    if format == "txt":
        line_break = '\n'
        filename = name + '.txt'
        fullname = os.path.join(path,filename)
        fixed_str = fullname.replace('\\', '//')  # 将 '\' 替换为 '//'
        # print('输出文件地址为：' + fullname)  # 输出结果
        sCommand = ['SelectTreeItem ("%s")'%tree,
            'With ASCIIExport',
            '.Reset',
            '.SetfileType "csv"',
            '.FileName ("%s")'%fixed_str,
            '.Execute',
            'End With']
        sCommand = line_break.join(sCommand)
        modelers.add_to_history('txt', sCommand)

# -----------------------------------------------------------------------------
# 小工具
# -----------------------------------------------------------------------------
# 读取指定行数据
def read_specific_row(csv_file, row_number):
    with open(csv_file, 'r') as file:
        reader = csv.reader(file)
        for index, row in enumerate(reader):
            if index == row_number:
                return row
            
def count_csv_files(folder_path, file_pattern):
    """计算给定文件夹中符合特定模式的CSV文件数量"""
    file_path_pattern = os.path.join(folder_path, file_pattern)
    return len(glob.glob(file_path_pattern))

# 保存 i 到文件
def save_progress(i, file_path):
    with open(file_path, 'w') as file:
        file.write(str(i))

# 从文件加载 i 的值
def load_progress(file_path):
    try:
        with open(file_path, 'r') as file:
            return int(file.read().strip())
    except FileNotFoundError:
        return 0  # 如果文件不存在，返回默认值

def crop_white_areas(image_path):
    image = Image.open(image_path)
    image_data = image.load()
    width, height = image.size

    top_crop, bottom_crop, left_crop, right_crop = 0, 0, 0, 0

    # 上边界裁剪
    for y in range(height):
        if all(image_data[x, y] == (255, 255, 255, 255) for x in range(width)):
            top_crop += 1
        else:
            break

    # 下边界裁剪
    for y in range(height-1, 0, -1):
        if all(image_data[x, y] == (255, 255, 255, 255) for x in range(width)):
            bottom_crop += 1
        else:
            break

    # 左边界裁剪
    for x in range(width):
        if all(image_data[x, y] == (255, 255, 255, 255) for y in range(height)):
            left_crop += 1
        else:
            break

    # 右边界裁剪
    for x in range(width-1, 0, -1):
        if all(image_data[x, y] == (255, 255, 255, 255) for y in range(height)):
            right_crop += 1
        else:
            break

    cropped_image = image.crop((left_crop, top_crop, width-right_crop, height-bottom_crop))
    cropped_image.save(image_path)  # 保存裁剪后的图片
    print('成功裁剪白色区域')

# -----------------------------------------------------------------------------
# 准备开始
# -----------------------------------------------------------------------------
# 调用函数加载 now
progress_file_path = os.path.join(folder_paths[3], "progress.txt")
now = load_progress(progress_file_path)
# 仿真总数量，每50个仿真在一个CST文件里（减少硬盘占用）
num_files = 1000
j_num = (num_files-now)//50

for j in range(0, j_num):
    filename = f'CST{now}_{now+50}.cst'#这里修改为仿真文件名称
    fullname = os.path.join(folder_paths[0],filename)
    # -----------------------------------------------------------------------------
    # 频率
    # -----------------------------------------------------------------------------
    Fre_start,Fre_stop = 2,2.5
    #仿真采样点数（默认1001）
    num = 500
    #二氧化饭参数
    name, value= "S",10
    de = DesignEnvironment()
    mws = de.new_mws()
    mws.save(path = fullname,allow_overwrite = True)
    modelers = mws.model3d
    modelers.add_to_history('StoreParameter','StoreParameter("%s","%f")' % (name,value))
    # 创建材料
    line_break = '\n'#换行
    sCommand = f'''With Material 
    .Reset 
    .Name "Au"
    .Folder ""
    .Rho "0.0"
    .ThermalType "Normal"
    .ThermalConductivity "0"
    .HeatCapacity "0"
    .DynamicViscosity "0"
    .Emissivity "0"
    .MetabolicRate "0.0"
    .VoxelConvection "0.0"
    .BloodFlow "0"
    .MechanicsType "Unused"
    .FrqType "all"
    .Type "Normal"
    .MaterialUnit "Frequency", "THz"
    .MaterialUnit "Geometry", "um"
    .MaterialUnit "Time", "ns"
    .MaterialUnit "Temperature", "Kelvin"
    .Epsilon "1"
    .Mu "1"
    .Sigma "4.5*10^7"
    .TanD "0.0"
    .TanDFreq "0.0"
    .TanDGiven "False"
    .TanDModel "ConstTanD"
    .EnableUserConstTanDModelOrderEps "False"
    .ConstTanDModelOrderEps "1"
    .SetElParametricConductivity "False"
    .ReferenceCoordSystem "Global"
    .CoordSystemType "Cartesian"
    .SigmaM "0"
    .TanDM "0.0"
    .TanDMFreq "0.0"
    .TanDMGiven "False"
    .TanDMModel "ConstTanD"
    .EnableUserConstTanDModelOrderMu "False"
    .ConstTanDModelOrderMu "1"
    .SetMagParametricConductivity "False"
    .DispModelEps  "None"
    .DispModelMu "None"
    .DispersiveFittingSchemeEps "Nth Order"
    .MaximalOrderNthModelFitEps "10"
    .ErrorLimitNthModelFitEps "0.1"
    .UseOnlyDataInSimFreqRangeNthModelEps "False"
    .DispersiveFittingSchemeMu "Nth Order"
    .MaximalOrderNthModelFitMu "10"
    .ErrorLimitNthModelFitMu "0.1"
    .UseOnlyDataInSimFreqRangeNthModelMu "False"
    .UseGeneralDispersionEps "False"
    .UseGeneralDispersionMu "False"
    .NonlinearMeasurementError "1e-1"
    .NLAnisotropy "False"
    .NLAStackingFactor "1"
    .NLADirectionX "1"
    .NLADirectionY "0"
    .NLADirectionZ "0"
    .Colour "1", "1", "0" 
    .Wireframe "False" 
    .Reflection "False" 
    .Allowoutline "True" 
    .Transparentoutline "False" 
    .Transparency "0" 
    .Create
    End With'''
    modelers.add_to_history('define_Au', sCommand)

    sCommand = f'''With Material
    .Reset
    .Name "VO2"
    .Folder ""
    .Rho "0.0"
    .ThermalType "Normal"
    .ThermalConductivity "0"
    .HeatCapacity "0"
    .DynamicViscosity "0"
    .Emissivity "0"
    .MetabolicRate "0.0"
    .VoxelConvection "0.0"
    .BloodFlow "0"
    .MechanicsType "Unused"
    .FrqType "all"
    .Type "Normal"
    .MaterialUnit "Frequency", "GHz"
    .MaterialUnit "Geometry", "um"
    .MaterialUnit "Time", "ns"
    .MaterialUnit "Temperature", "Kelvin"
    .Epsilon "1"
    .Mu "1"
    .Sigma "0.0"
    .TanD "0.0"
    .TanDFreq "0.0"
    .TanDGiven "False"
    .TanDModel "ConstTanD"
    .EnableUserConstTanDModelOrderEps "False"
    .ConstTanDModelOrderEps "1"
    .SetElParametricConductivity "False"
    .ReferenceCoordSystem "Global"
    .CoordSystemType "Cartesian"
    .SigmaM "0.0"
    .TanDM "0.0"
    .TanDMFreq "0.0"
    .TanDMGiven "False"
    .TanDMModel "ConstTanD"
    .EnableUserConstTanDModelOrderMu "False"
    .ConstTanDModelOrderMu "1"
    .SetMagParametricConductivity "False"
    .DispModelEps  "Drude"
    .EpsInfinity "12"
    .DispCoeff1Eps "sqr(S/300000)*1.4*10^15"
    .DispCoeff2Eps "(1/(2*pi))*5.75*10^13"
    .DispCoeff3Eps "0.0"
    .DispCoeff4Eps "0.0"
    .DispModelMu "None"
    .DispersiveFittingSchemeEps "Nth Order"
    .MaximalOrderNthModelFitEps "10"
    .ErrorLimitNthModelFitEps "0.1"
    .DispersiveFittingSchemeMu "Nth Order"
    .MaximalOrderNthModelFitMu "10"
    .ErrorLimitNthModelFitMu "0.1"
    .UseGeneralDispersionEps "False"
    .UseGeneralDispersionMu "False"
    .NonlinearMeasurementError "1e-1"
    .NLAnisotropy "False"
    .NLAStackingFactor "1"
    .NLADirectionX "1"
    .NLADirectionY "0"
    .NLADirectionZ "0"
    .Colour "0", "0", "1" 
    .Wireframe "False" 
    .Reflection "False" 
    .Allowoutline "True" 
    .Transparentoutline "False" 
    .Transparency "0" 
    .Create
    End With'''

    modelers.add_to_history('define_VO2', sCommand)
    
    sCommand = f'''With Material 
        .Reset 
        .Name "PI"
        .Folder ""
        .Rho "0.0"
        .ThermalType "Normal"
        .ThermalConductivity "0"
        .HeatCapacity "0"
        .DynamicViscosity "0"
        .Emissivity "0"
        .MetabolicRate "0.0"
        .VoxelConvection "0.0"
        .BloodFlow "0"
        .MechanicsType "Unused"
        .FrqType "all"
        .Type "Normal"
        .MaterialUnit "Frequency", "THz"
        .MaterialUnit "Geometry", "um"
        .MaterialUnit "Time", "ns"
        .MaterialUnit "Temperature", "Kelvin"
        .Epsilon "4.3"
        .Mu "1"
        .Sigma "1.11237"
        .TanD "0.00093"
        .TanDFreq "1"
        .TanDGiven "True"
        .TanDModel "ConstSigma"
        .EnableUserConstTanDModelOrderEps "False"
        .ConstTanDModelOrderEps "1"
        .SetElParametricConductivity "False"
        .ReferenceCoordSystem "Global"
        .CoordSystemType "Cartesian"
        .SigmaM "0"
        .TanDM "0.0"
        .TanDMFreq "0.0"
        .TanDMGiven "False"
        .TanDMModel "ConstTanD"
        .EnableUserConstTanDModelOrderMu "False"
        .ConstTanDModelOrderMu "1"
        .SetMagParametricConductivity "False"
        .DispModelEps "None"
        .DispModelMu "None"
        .DispersiveFittingSchemeEps "Nth Order"
        .MaximalOrderNthModelFitEps "10"
        .ErrorLimitNthModelFitEps "0.1"
        .UseOnlyDataInSimFreqRangeNthModelEps "False"
        .DispersiveFittingSchemeMu "Nth Order"
        .MaximalOrderNthModelFitMu "10"
        .ErrorLimitNthModelFitMu "0.1"
        .UseOnlyDataInSimFreqRangeNthModelMu "False"
        .UseGeneralDispersionEps "False"
        .UseGeneralDispersionMu "False"
        .NonlinearMeasurementError "1e-1"
        .NLAnisotropy "False"
        .NLAStackingFactor "1"
        .NLADirectionX "1"
        .NLADirectionY "0"
        .NLADirectionZ "0"
        .Colour "0", "0.501961", "0.752941" 
        .Wireframe "False" 
        .Reflection "False" 
        .Allowoutline "True" 
        .Transparentoutline "False" 
        .Transparency "0" 
        .Create
    End With'''

    modelers.add_to_history('define_PI', sCommand)

    sCommand = f'''With Material 
    .Reset 
    .Name "SiO2"
    .Folder ""
    .Rho "0.0"
    .ThermalType "Normal"
    .ThermalConductivity "0"
    .HeatCapacity "0"
    .DynamicViscosity "0"
    .Emissivity "0"
    .MetabolicRate "0.0"
    .VoxelConvection "0.0"
    .BloodFlow "0"
    .MechanicsType "Unused"
    .FrqType "all"
    .Type "Normal"
    .MaterialUnit "Frequency", "THz"
    .MaterialUnit "Geometry", "um"
    .MaterialUnit "Time", "ns"
    .MaterialUnit "Temperature", "Kelvin"
    .Epsilon "3.75"
    .Mu "1"
    .Sigma "0"
    .TanD "0.0"
    .TanDFreq "0.0"
    .TanDGiven "False"
    .TanDModel "ConstTanD"
    .EnableUserConstTanDModelOrderEps "False"
    .ConstTanDModelOrderEps "1"
    .SetElParametricConductivity "False"
    .ReferenceCoordSystem "Global"
    .CoordSystemType "Cartesian"
    .SigmaM "0"
    .TanDM "0.0"
    .TanDMFreq "0.0"
    .TanDMGiven "False"
    .TanDMModel "ConstTanD"
    .EnableUserConstTanDModelOrderMu "False"
    .ConstTanDModelOrderMu "1"
    .SetMagParametricConductivity "False"
    .DispModelEps  "None"
    .DispModelMu "None"
    .DispersiveFittingSchemeEps "Nth Order"
    .MaximalOrderNthModelFitEps "10"
    .ErrorLimitNthModelFitEps "0.1"
    .UseOnlyDataInSimFreqRangeNthModelEps "False"
    .DispersiveFittingSchemeMu "Nth Order"
    .MaximalOrderNthModelFitMu "10"
    .ErrorLimitNthModelFitMu "0.1"
    .UseOnlyDataInSimFreqRangeNthModelMu "False"
    .UseGeneralDispersionEps "False"
    .UseGeneralDispersionMu "False"
    .NonlinearMeasurementError "1e-1"
    .NLAnisotropy "False"
    .NLAStackingFactor "1"
    .NLADirectionX "1"
    .NLADirectionY "0"
    .NLADirectionZ "0"
    .Colour "1", "1", "0" 
    .Wireframe "False" 
    .Reflection "False" 
    .Allowoutline "True" 
    .Transparentoutline "False" 
    .Transparency "0" 
    .Create
    End With
    '''
    modelers.add_to_history('define_SiO2', sCommand)
    line_break = '\n'#换行
    sCommand = f'''With Material 
    .Reset 
    .Name "PSi"
    .Folder ""
    .Rho "0.0"
    .ThermalType "Normal"
    .ThermalConductivity "0"
    .HeatCapacity "0"
    .DynamicViscosity "0"
    .Emissivity "0"
    .MetabolicRate "0.0"
    .VoxelConvection "0.0"
    .BloodFlow "0"
    .MechanicsType "Unused"
    .FrqType "all"
    .Type "Normal"
    .MaterialUnit "Frequency", "THz"
    .MaterialUnit "Geometry", "um"
    .MaterialUnit "Time", "ns"
    .MaterialUnit "Temperature", "Kelvin"
    .Epsilon "11.7"
    .Mu "1"
    .Sigma "1.5*10^5"
    .TanD "0.0"
    .TanDFreq "0.0"
    .TanDGiven "False"
    .TanDModel "ConstTanD"
    .EnableUserConstTanDModelOrderEps "False"
    .ConstTanDModelOrderEps "1"
    .SetElParametricConductivity "False"
    .ReferenceCoordSystem "Global"
    .CoordSystemType "Cartesian"
    .SigmaM "0"
    .TanDM "0.0"
    .TanDMFreq "0.0"
    .TanDMGiven "False"
    .TanDMModel "ConstTanD"
    .EnableUserConstTanDModelOrderMu "False"
    .ConstTanDModelOrderMu "1"
    .SetMagParametricConductivity "False"
    .DispModelEps  "None"
    .DispModelMu "None"
    .DispersiveFittingSchemeEps "Nth Order"
    .MaximalOrderNthModelFitEps "10"
    .ErrorLimitNthModelFitEps "0.1"
    .UseOnlyDataInSimFreqRangeNthModelEps "False"
    .DispersiveFittingSchemeMu "Nth Order"
    .MaximalOrderNthModelFitMu "10"
    .ErrorLimitNthModelFitMu "0.1"
    .UseOnlyDataInSimFreqRangeNthModelMu "False"
    .UseGeneralDispersionEps "False"
    .UseGeneralDispersionMu "False"
    .NonlinearMeasurementError "1e-1"
    .NLAnisotropy "False"
    .NLAStackingFactor "1"
    .NLADirectionX "1"
    .NLADirectionY "0"
    .NLADirectionZ "0"
    .Colour "0.752941", "0.752941", "0.752941" 
    .Wireframe "False" 
    .Reflection "False" 
    .Allowoutline "True" 
    .Transparentoutline "False" 
    .Transparency "0" 
    .Create
    End With'''

    modelers.add_to_history('define_PSi', sCommand)
    # 设置单位
    sCommand = '''With Units
    .Geometry "um"
    .Frequency "THz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "H"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "F"
    End With'''

    # 设置环境温度
    ambient_temp_command = 'ThermalSolver.AmbientTemperature "0"'
    # 设置频率范围
    freq_range_command = 'Solver.FrequencyRange "%f","%f"'  % (Fre_start,Fre_stop)
    # 绘制盒子
    draw_box_command = 'Plot.DrawBox False'
    # 设置背景属性
    background_command = '''With Background
    .Type "Normal"
    .Epsilon "1.0"
    .Mu "1.0"
    .Rho "1.204"
    .ThermalType "Normal"
    .ThermalConductivity "0.026"
    .HeatCapacity "1.005"
    .XminSpace "0.0"
    .XmaxSpace "0.0"
    .YminSpace "0.0"
    .YmaxSpace "0.0"
    .ZminSpace "0.0"
    .ZmaxSpace "0.0"
    End With'''

    # 设置Floquet端口边界
    floquet_port_command = '''With FloquetPort
    .Reset
    .SetDialogTheta "0"
    .SetDialogPhi "0"
    .SetSortCode "+beta/pw"
    .SetCustomizedListFlag "False"
    .Port "Zmin"
    .SetNumberOfModesConsidered "2"
    .Port "Zmax"
    .SetNumberOfModesConsidered "2"
    End With'''
    # 确保参数存在并设置描述
    parameter_commands = [
            'MakeSureParameterExists "theta", "0"',
            'SetParameterDescription "theta", "spherical angle of incident plane wave"',
            'MakeSureParameterExists "phi", "0"',
            'SetParameterDescription "phi", "spherical angle of incident plane wave"'
    ]
    parameter_command = line_break.join(parameter_commands)
    # 边界条件的设置
    boundary_command = '''With Boundary
    .Xmin "unit cell"
    .Xmax "unit cell"
    .Ymin "unit cell"
    .Ymax "unit cell"
    .Zmin "expanded open"
    .Zmax "expanded open"
    .Xsymmetry "none"
    .Ysymmetry "none"
    .Zsymmetry "none"
    .XPeriodicShift "0.0"
    .YPeriodicShift "0.0"
    .ZPeriodicShift "0.0"
    .PeriodicUseConstantAngles "False"
    .SetPeriodicBoundaryAngles "theta", "phi"
    .SetPeriodicBoundaryAnglesDirection "inward"
    .UnitCellFitToBoundingBox "True"
    .UnitCellDs1 "0.0"
    .UnitCellDs2 "0.0"
    .UnitCellAngle "90.0"
    End With'''
    # 设置网格
    mesh_command = '''With MeshSettings
        .SetMeshType "Tet"
        .Set "Version", 1%
    End With'''

    # FDSolver设置
    fdsolver_command = '''With FDSolver
        .Reset 
        .SetMethod "Tetrahedral", "General purpose" 
        .OrderTet "Second" 
        .OrderSrf "First" 
        .Stimulation "List", "List" 
        .ResetExcitationList 
        .AddToExcitationList "Zmax", "TE(0,0);TM(0,0)" 
        .AutoNormImpedance "False" 
        .NormingImpedance "50" 
        .ModesOnly "False" 
        .ConsiderPortLossesTet "True" 
        .SetShieldAllPorts "False" 
        .AccuracyHex "1e-6" 
        .AccuracyTet "1e-4" 
        .AccuracySrf "1e-3" 
        .LimitIterations "False" 
        .MaxIterations "0" 
        .SetCalcBlockExcitationsInParallel "True", "True", "" 
        .StoreAllResults "False" 
        .StoreResultsInCache "False" 
        .UseHelmholtzEquation "True" 
        .LowFrequencyStabilization "False" 
        .Type "Auto" 
        .MeshAdaptionHex "False" 
        .MeshAdaptionTet "True" 
        .AcceleratedRestart "True" 
        .FreqDistAdaptMode "Distributed" 
        .NewIterativeSolver "True" 
        .TDCompatibleMaterials "False" 
        .ExtrudeOpenBC "False" 
        .SetOpenBCTypeHex "Default" 
        .SetOpenBCTypeTet "Default" 
        .AddMonitorSamples "True" 
        .CalcPowerLoss "True" 
        .CalcPowerLossPerComponent "False" 
        .StoreSolutionCoefficients "True" 
        .UseDoublePrecision "False" 
        .UseDoublePrecision_ML "True" 
        .MixedOrderSrf "False" 
        .MixedOrderTet "False" 
        .PreconditionerAccuracyIntEq "0.15" 
        .MLFMMAccuracy "Default" 
        .MinMLFMMBoxSize "0.3" 
        .UseCFIEForCPECIntEq "True" 
        .UseEnhancedCFIE2 "True" 
        .UseFastRCSSweepIntEq "false" 
        .UseSensitivityAnalysis "False" 
        .UseEnhancedNFSImprint "False" 
        .RemoveAllStopCriteria "Hex"
        .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
        .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
        .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
        .RemoveAllStopCriteria "Tet"
        .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
        .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
        .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
        .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
        .RemoveAllStopCriteria "Srf"
        .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
        .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
        .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
        .SweepMinimumSamples "3" 
        .SetNumberOfResultDataSamples "1001" 
        .SetResultDataSamplingMode "Automatic" 
        .SweepWeightEvanescent "1.0" 
        .AccuracyROM "1e-4" 
        .AddSampleInterval "", "", "1", "Automatic", "True" 
        .AddSampleInterval "", "", "", "Automatic", "False" 
        .MPIParallelization "False"
        .UseDistributedComputing "False"
        .NetworkComputingStrategy "RunRemote"
        .NetworkComputingJobCount "3"
        .UseParallelization "True"
        .MaxCPUs "96"
        .MaximumNumberOfCPUDevices "2"
    End With'''

    # IESolver设置
    iesolver_command = '''With IESolver
        .Reset 
        .UseFastFrequencySweep "True" 
        .UseIEGroundPlane "False" 
        .SetRealGroundMaterialName "" 
        .CalcFarFieldInRealGround "False" 
        .RealGroundModelType "Auto" 
        .PreconditionerType "Auto" 
        .ExtendThinWireModelByWireNubs "False" 
        .ExtraPreconditioning "False" 
    End With'''
    # IESolver的进一步设置
    iesolver_additional_command = '''With IESolver
        .SetFMMFFCalcStopLevel "0" 
        .SetFMMFFCalcNumInterpPoints "6" 
        .UseFMMFarfieldCalc "True" 
        .SetCFIEAlpha "0.500000" 
        .LowFrequencyStabilization "False" 
        .LowFrequencyStabilizationML "True" 
        .Multilayer "False" 
        .SetiMoMACC_I "0.0001" 
        .SetiMoMACC_M "0.0001" 
        .DeembedExternalPorts "True" 
        .SetOpenBC_XY "True" 
        .OldRCSSweepDefintion "False" 
        .SetRCSOptimizationProperties "True", "100", "0.00001" 
        .SetAccuracySetting "Custom" 
        .CalculateSParaforFieldsources "True" 
        .ModeTrackingCMA "True" 
        .NumberOfModesCMA "3" 
        .StartFrequencyCMA "-1.0" 
        .SetAccuracySettingCMA "Default" 
        .FrequencySamplesCMA "0" 
        .SetMemSettingCMA "Auto" 
        .CalculateModalWeightingCoefficientsCMA "True" 
        .DetectThinDielectrics "True" 
    End With'''

    mesh_settings_commands = [
            'With MeshSettings',
            '.SetMeshType "Tet"',
            '.Set "Version", 1',
            'End With'
    ]
    mesh_settings_command = line_break.join(mesh_settings_commands)

    # 更改求解器类型为高频频域（HF Frequency Domain）
    change_solver_type_command = 'ChangeSolverType("HF Frequency Domain")'
    modelers.add_to_history('SetAllowFloatDirectSolver','FDSolver.SetAllowFloatDirectSolver "True"')
    modelers.add_to_history('SetKeepSolutionCoefficients','FDSolver.SetKeepSolutionCoefficients "All"')

    # 将所有命令添加到模型的历史记录
    modelers.add_to_history('Units', sCommand)
    modelers.add_to_history('Ambient Temperature', ambient_temp_command)
    modelers.add_to_history('Frequency Range', freq_range_command)
    modelers.add_to_history('Draw Box', draw_box_command)
    modelers.add_to_history('Background', background_command)
    modelers.add_to_history('Floquet Port', floquet_port_command)
    modelers.add_to_history('Parameter', parameter_command)
    modelers.add_to_history('Boundary', boundary_command)
    modelers.add_to_history('Mesh', mesh_command)
    modelers.add_to_history('FDSolver', fdsolver_command)
    modelers.add_to_history('IESolver', iesolver_command)
    modelers.add_to_history('IESolver Additional', iesolver_additional_command)
    modelers.add_to_history('MeshSettings', mesh_settings_command)
    modelers.add_to_history('Change Solver Type', change_solver_type_command)
    # 基础设置
    ChangeColour('VO2',1,0,0)
    ChangeColour('PSi',"0.752941", "0.752941", "0.752941")

    first_iteration = True
    for i in range(now, now+50):
        if first_iteration:
            windows_now = now
            first_iteration = False
        print(f"当前进度为{i}/{num_files-1}")
        # -----------------------------------------------------------------------------
        # 导入生成数据集
        # -----------------------------------------------------------------------------
        csv_file ='D:\学习\cst测试\咸鱼\data_set.csv'
        row_number = i + 1  # 读取第3行数据，索引从0开始
        specific_row = read_specific_row(csv_file, row_number)
        Px , Py = 100 , 100
        numeric_data = [float(item) for item in specific_row]
        l4, l3, l2, l1, d1, d2, h = numeric_data
        t1 , t2 = 0.2 , 0.25
        #建模
        brick('1','component1','Au',[f'-{Px}/2',f'{Px}/2'],[f'-{Py}/2',f'{Py}/2'],["0",f"{t2}"])
        wcs_face('component1','1', faceid=1)
        brick('2','component1','PI',[f'-{Px}/2',f'{Px}/2'],[f'-{Py}/2',f'{Py}/2'],["0",f"{d2}"])
        wcs_face('component1','2', faceid=1)
        brick('3','component1','VO2',[f'-{Px}/2',f'{Px}/2'],[f'-{Py}/2',f'{Py}/2'],["0",f"{h}"])
        wcs_face('component1','3', faceid=1)
        brick('4','component1','PI',[f'-{Px}/2',f'{Px}/2'],[f'-{Py}/2',f'{Py}/2'],["0",f"{d1}"])
        wcs_face('component1','4', faceid=1)
        brick('5_1','component1','PSi',[f'-{l1}/2',f'{l1}/2'],[f'-{l1}/2',f'{l1}/2'],["0",f"{t1}"])
        brick('5_11','component1','PSi',[f'-{l2}/2',f'{l2}/2'],[f'-{l2}/2',f'{l2}/2'],["0",f"{t1}"])
        modelers.add_to_history('Subtract51','Solid.Subtract "component1:5_1", "component1:5_11"')
        brick('5_2','component1','PSi',[f'-{l3}/2',f'{l3}/2'],[f'-{l3}/2',f'{l3}/2'],["0",f"{t1}"])
        brick('5_22','component1','PSi',[f'-{l4}/2',f'{l4}/2'],[f'-{l4}/2',f'{l4}/2'],["0",f"{t1}"])
        modelers.add_to_history('Subtract52','Solid.Subtract "component1:5_2", "component1:5_22"')
        
        windows = gw.getWindowsWithTitle(f'CST{windows_now}_{windows_now+50} - CST Studio Suite 2024')
        if windows:
            # 选择第一个匹配的窗口
            window = windows[0]
            # 最大化窗口
            window.maximize()
        else:
            print("未找到指定标题的窗口")

        #保存图片            
        modelers.add_to_history('ResetZoom','Plot.SetGradientBackground"False"')
        modelers.add_to_history('ResetZoom','Wcs.ActivateWCS"global"')
        modelers.add_to_history('ResetZoom','Plot.ZoomToStructure')
        modelers.add_to_history('StoreParameter','Plot.RestoreView"Perspective"')
        modelers.add_to_history('DrawBox','Plot.DrawBox"False"')
        modelers.add_to_history('DrawWorkplane','Plot.DrawWorkplane "False"')
        plot_file_path = os.path.join(folder_paths[2], f"Plot{i}.png")
        modelers.add_to_history('DrawWorkplane',f'Plot.ExportImage ("{plot_file_path}", 800, 800)')
        crop_white_areas(plot_file_path)  # 替换为你的图像文件路径

        #运行仿真
        modelers.run_solver()        
        exdata(sp="SZmax(1),Zmax(1)",type="mag",format="txt",path=folder_paths[1],name=f"output{i}")        
        modelers.add_to_history('detresult','DeleteResults')
        modelers.add_to_history('ComponentDelete','Component.Delete "component1" ')
        now=i
        save_progress(now, progress_file_path)
    # mws.close()
    de.close()
    
# HVDC_2StageRO_CCG
这个仓库包含了针对高压直流潮流问题(HVDC)的二阶段鲁棒优化问题的pythpn求解代码。使用了行列生成算法(C&CG)进行求解。
This repository contains Python code for solving a two-stage robust optimization problem related to High Voltage Direct Current (HVDC) power flow. The solution employs the Column and Constraint Generation (C&CG) algorithm.

## Key Features
1. Establishes a robust optimization model for HVDC, including the objective function and constraints.
2. Utilizes the Column and Constraint Generation (C&CG) algorithm to solve the optimization problem.
3. Analyzes the solution results and generates relevant data and visualizations.

## Tech Stack
- Python
- Jupyter Notebook

## License
Not specified

## 模型建立
（具体内容详见"模型.pdf"）

模型包括一条高压直流传输线，送端包含一台风电与五台火电机组，受端包含五台火电机组与负荷。具体数据详见"P10data.xlsx"。

目标函数：最小化成本，成本包含火电启停成本，火电出力成本，弃风成本。

约束条件：包含火电运行约束，直流潮流运行约束，功率平衡约束等。

不确定集：设置为负荷在24小时的波动范围[-1,1]。

## 求解算法
（详见main.py）

采用行列生成算法（C&CG）进行求解。其中，上层规划问题为机组启停与直流潮流变化，下层规划问题为风火出力传输功率与直流潮流传输功率。

## 结果分析
采用"Data.xlsx"的调度结果如下图：

![最坏场景下调度结果](https://github.com/LuminaryEdge/HVDC_2StageRO_CCG/assets/120100760/d75cc265-604f-4843-8b49-a38beef3d762)



# HoldExcel

- [HoldExcel](#holdexcel)
  - [从 Excel 中读取数据的方式](#%e4%bb%8e-excel-%e4%b8%ad%e8%af%bb%e5%8f%96%e6%95%b0%e6%8d%ae%e7%9a%84%e6%96%b9%e5%bc%8f)
    - [字典可重模式](#%e5%ad%97%e5%85%b8%e5%8f%af%e9%87%8d%e6%a8%a1%e5%bc%8f)
    - [字典不可重模式](#%e5%ad%97%e5%85%b8%e4%b8%8d%e5%8f%af%e9%87%8d%e6%a8%a1%e5%bc%8f)
    - [列表模式](#%e5%88%97%e8%a1%a8%e6%a8%a1%e5%bc%8f)
  - [关于 Hard Code](#%e5%85%b3%e4%ba%8e-hard-code)

学习使用 `openpyxl` 包操作 `excel` 文件。

## 从 Excel 中读取数据的方式

### 字典可重模式

输入数据可以为任意列数，第一行是标题行，剩余行为数据。其中第一列是 key，其他列是数据，输出结果为 python 的 字典对象(dictionary)，第一列作为字典的 key 值，value 为 list，内容是由第二列和第三列组成的字典。

**原格式样式**

| ConfigID | ColumnName1 | ColumnName2 |
| -------: | :---------: | :---------: |
|        1 |      A      |      F      |
|        1 |      B      |     SUM     |
|        2 |      C      |      D      |
|        3 |      F      |      W      |

**输出-Dictionary**

```python
{
    1: [
        {'ColumnName1': 'A', 'ColumnName2': 'F'}, 
        {'ColumnName1': 'B', 'ColumnName2': 'SUM'}
    ],
    2: [
        {'ColumnName1': 'C', 'ColumnName2': 'D'}
    ],
    3: [
        {'ColumnName1': 'F', 'ColumnName2': 'W'}
    ]
}
```

### 字典不可重模式

输入数据可以为任意列数，第一行是标题行，剩余行为数据。其中第一列是 key，要求 key 值不可重复，其他列是数据，输出结果为 python 的 字典对象(dictionary)，第一列作为字典的 key 值，value 为由其他列组成的字典。

**原格式样式**

| ConfigID | ColumnName1 | ColumnName2 | ColumnName3 |
| -------: | :---------: | :---------: | :---------: |
|        1 |   5A.xlsx   |    预算     |    50.00    |
|        2 |   7A.xlsx   |   5年规划   |     100     |
|        3 |   9B.xlsx   |  维护预算   |     500     |

**输出-Dictionary**

```python
{
    1: {
        'ColumnName1': '5A.xlsx',
        'ColumnName2': '预算',
        'ColumnName3': '50.00'
    },
    2: {
        'ColumnName1': '7A.xlsx',
        'ColumnName2': '5年规划',
        'ColumnName3': '100'
    },
    3: {
        'ColumnName1': '9B.xlsx',
        'ColumnName2': '维护预算',
        'ColumnName3': '500'
    }
}
```

### 列表模式

输入数据可以为任意列数，第一行是标题行，剩余行为数据。其中第一列是 key，其他列是数据，输出结果为 python 的列表对象(list)，每个元素代表一行数据，数据由其他列组成的字典。

**原格式样式**

| ConfigID | ColumnName1 | ColumnName2 | ColumnName3 |
| -------: | ----------: | ----------: | ----------: |
|        1 |         100 |         200 |         300 |
|        1 |        1000 |        2000 |        3000 |
|        2 |       10000 |       20000 |       30000 |

**输出-Dictionary**

```python
[
    {
        'ConfigID': '1', 'ColumnName1': '100',
        'ColumnName2': '200', 'ColumnName3': '300'
    },
    {
        'ConfigID': '1', 'ColumnName1': '1000',
        'ColumnName2': '2000', 'ColumnName3': '3000'
    },
    {
        'ConfigID': '2', 'ColumnName1': '10000',
        'ColumnName2': '20000', 'ColumnName3': '30000'
    },
]
```

## 关于 Hard Code

为了避免出现 Hard Code 问题，将涉及用户可能修改的配置数据放置在 `config.yml` 配置文件中，系统内部需要使用的配置数据放置在 `global_var.py` 文件中。

整个系统中允许在 `global_var.py` 文件中写入 Hard Code。
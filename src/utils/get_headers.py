pl_head_dict = {'2024': {'Profit Center': 'Profit Center',
                         "Total Direct Costs": "Total Direct     Costs",
                         "Total Operating Costs": "Total Operating Costs"},
                '2025': {'Profit Center': 'Профит центр',
                         "Total Direct Costs": "Итого прямые расходы",
                         "Total Operating Costs": "Итого операционные расходы"}
                }
secured_rev_head_dict = {'2024': {'Analytics': 'Analytics'},
                         '2025': {'Analytics': 'Аналитика данных'},
                         }
"""
Profit Center', 'Компания', 'Номер контракта',
                             'Сумма контракта без НДС', 'Дата начала контракта',
                             'Дата завершения контракта'
                             

"""


def pl_header(year: float, month: float):
    if year <= 2024 or (year == 2025 and month < 5):
        return pl_head_dict['2024']
    else:
        return pl_head_dict['2025']


def secured_rev_header(year: float, month: float):
    if year <= 2024 or (year == 2025 and month < 5):
        return secured_rev_head_dict['2024']
    else:
        return secured_rev_head_dict['2025']

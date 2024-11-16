patterns_to_remove = [
    r"Order number: \d+",
    r", \d+/\d+/\d+ \d+:\d+-\d+:\d+",
    r"Cost of goods £\d+.\d+",
    r"Offers savings -£\d+.\d+",
    r"Card: \w+",
    r"Last four digits: \d+",
    r"You have paid VAT of £\d+.\d+ on VATable items totalling £\d+.\d+",
    r"Offers savings.*",
    r"\(£\d+.\d+/EACH\)",
    r"\(£\d+.\d+/ EACH\)",
    r"8ZL\d+-\d+",
    r"Vouchers and extras -£\d+.\d+",
]
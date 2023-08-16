import os
import pandas as pd
import stocks
from datetime import datetime

def main():
    options, dividends = get_data()

    create_qif_file(options, dividends)


def get_data():
    options, dividends = stocks.getRobinhoodTransactions()
    options = options.loc[options['state'] == 'filled']
    dividends = dividends.loc[dividends['state'] == 'paid']


    options['last_transaction_at'] = pd.to_datetime(options['last_transaction_at'])
    options = options.astype({
        'price': float,
        'quantity': float,
    })


    dividends['payable_date'] = pd.to_datetime(dividends['payable_date'], utc=True)
    dividends = dividends.astype({
        'amount': float,
    })


    date_input = input("Get Transaction After (yyyy-mm-dd): ")
    date_input = pd.Timestamp(date_input, tz='UTC')

    options = options.loc[options['last_transaction_at'] >= date_input]
    dividends = dividends.loc[dividends['payable_date'] >= date_input]
    return options, dividends


def get_quicken_name(robinhood_name: str):
    file_name = 'stock_mapping_robinhood_quicken.csv'

    if not os.path.exists(file_name):
        mapping_df = pd.DataFrame([], columns=['Robinhood Name', 'Quicken Name'])
    else:
        mapping_df = pd.read_csv(file_name)

    quicken_name = mapping_df.loc[mapping_df['Robinhood Name'] == robinhood_name, 'Quicken Name']

    if len(quicken_name) > 1:
        raise ValueError(f"Multiple Values found for {robinhood_name}. Update the {file_name} to fix this")
    elif len(quicken_name) == 0:
        print(f"Could not map Robinhood Stock: {robinhood_name}")
        quicken_name = input("Please input the matching quicken name: ").strip()
        new_df = pd.DataFrame([[robinhood_name, quicken_name]], columns=['Robinhood Name', 'Quicken Name'])
        mapping_df = pd.concat([mapping_df, new_df])

        mapping_df.to_csv(file_name, index=False)
    else:
        quicken_name = quicken_name.iloc[0]

    return quicken_name

def create_qif_file(options: pd.DataFrame, dividends: pd.DataFrame):
    qif_lines = ['!Type:Invst']


    for _, transaction in options.iterrows():
        year = transaction['last_transaction_at'].year
        day = transaction['last_transaction_at'].day
        month = transaction['last_transaction_at'].month

        quicken_name = get_quicken_name(transaction['Name'])        

        if day < 10:
            day = ' ' + str(day)
        else:
            day = str(day)
        year = str(year)[2:]

        qif_lines.append(f'D{month}/{day}\'{year}')
        if transaction['state'] == 'buy':
            qif_lines.append('NBuy')
        elif transaction['state'] == 'sell':
            qif_lines.append('NSell')

        # input("here:")
        price = transaction['price']
        quantity = transaction['quantity']
        net_gain = round(price * quantity, 2)
        qif_lines.append(f"Y{quicken_name}")
        qif_lines.append(f"I{price}")
        qif_lines.append(f"Q{quantity}")
        qif_lines.append(f"U{net_gain}")
        qif_lines.append(f"T{net_gain}")

        qif_lines.append("^")

    
    for _, dividend in dividends.iterrows():
        year = dividend['payable_date'].year
        day = dividend['payable_date'].day
        month = dividend['payable_date'].month

        quicken_name = get_quicken_name(dividend['Name'])        


        if day < 10:
            day = ' ' + str(day)
        else:
            day = str(day)
        year = str(year)[2:]

        qif_lines.append(f'D{month}/{day}\'{year}')
        qif_lines.append('NDiv')


        qif_lines.append(f"Y{quicken_name}")
        qif_lines.append(f"U{dividend['amount']}")
        qif_lines.append(f"T{dividend['amount']}")
        qif_lines.append("^")


    with open('output_quicken_file.qif', 'w') as f:
        f.write('\n'.join(qif_lines))

if __name__ == "__main__":
    main()
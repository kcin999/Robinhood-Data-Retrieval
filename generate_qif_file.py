"""Generates a QIF File which can be imported into Quicken"""
import os
from datetime import date
import pandas as pd
import stocks


def main():
    """Main Function"""
    options, dividends = get_data()

    create_qif_file(options, dividends)


def get_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Gets Options and divident data after the provided user data

    :return: Tuple consisting of a option and a dividend dataframe
    :rtype: tuple[pd.DataFrame, pd.DataFrame]
    """
    options, dividends = stocks.get_robinhood_transactions()
    options = options.loc[options['state'] == 'filled']
    dividends = dividends.loc[dividends['state'] == 'paid']

    options['last_transaction_at'] = pd.to_datetime(
        options['last_transaction_at'])
    options = options.astype({
        'price': float,
        'quantity': float,
    })

    dividends['payable_date'] = pd.to_datetime(
        dividends['payable_date'], utc=True)
    dividends = dividends.astype({
        'amount': float,
    })

    date_input = input("Get Transaction After (yyyy-mm-dd): ")
    date_input = pd.Timestamp(date_input, tz='UTC')

    options = options.loc[options['last_transaction_at'] > date_input]
    dividends = dividends.loc[dividends['payable_date'] > date_input]
    return options, dividends


def get_date_input() -> pd.Timestamp:
    """Gets the date input from the given file

    :return: Timestamp
    :rtype: pd.Timestamp
    """
    df = pd.read_csv('date_file.csv') 
    df['Date'] = pd.to_datetime(df['Date'])
    date_input = df.at[0,'Date']

    date_input = pd.Timestamp(date_input, tz="UTC")

    df.at[0, 'Date'] = date.today()
    df.to_csv('date_file.csv', index=False)

    return date_input


def get_quicken_name(robinhood_name: str) -> str:
    """Quicken names do not match what is robinhood. 
    This is a function that will help identify any names that do or do not exist

    :param robinhood_name: Robinhood Stock name
    :type robinhood_name: str
    :raises ValueError: Thrown if multiple quicken values are found for a given robinhood name
    :return: Matching Quicken Name
    :rtype: str
    """
    file_name = 'stock_mapping_robinhood_quicken.csv'

    if not os.path.exists(file_name):
        mapping_df = pd.DataFrame(
            [], columns=['Robinhood Name', 'Quicken Name'])
    else:
        mapping_df = pd.read_csv(file_name)

    quicken_name = mapping_df.loc[mapping_df['Robinhood Name']
                                  == robinhood_name, 'Quicken Name']

    if len(quicken_name) > 1:
        raise ValueError(
            f"Multiple Values found for {robinhood_name}. Update the {file_name} to fix this")
    elif len(quicken_name) == 0:
        print(f"Could not map Robinhood Stock: {robinhood_name}")
        quicken_name = input(
            "Please input the matching quicken name: ").strip()
        new_df = pd.DataFrame([[robinhood_name, quicken_name]], columns=[
                              'Robinhood Name', 'Quicken Name'])
        mapping_df = pd.concat([mapping_df, new_df])

        mapping_df.to_csv(file_name, index=False)
    else:
        quicken_name = quicken_name.iloc[0]

    return quicken_name


def create_qif_file(options: pd.DataFrame, dividends: pd.DataFrame):
    """Creates the QIF File which can then be imported into quicken

    :param options: Options DataFrame
    :type options: pd.DataFrame
    :param dividends: Dividends DataFrame
    :type dividends: pd.DataFrame
    """
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
        if transaction['side'] == 'buy' and transaction['state'] == 'filled':
            qif_lines.append('NBuy')
        elif transaction['side'] == 'sell' and transaction['state'] == 'filled':
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

    with open('output_quicken_file.qif', 'w') as file:
        file.write('\n'.join(qif_lines))


if __name__ == "__main__":
    main()

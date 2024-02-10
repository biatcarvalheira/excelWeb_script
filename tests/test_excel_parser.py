import re

# Your data
data = '''
    df['option Expiration date'] = option_expiration_date
    df['Strike'] = strike
    df['underlying symbol'] = underlying_symbol
    df['Type'] = stock_type
    df['mkt beta* mkt px*contracts'] = mkt_beta_px_contracts
    df['Qty'] = quantity
    df['mkt price *number of contracts'] = mkt_price_of_contracts
    df['Trade Price/premium'] = price
    df['trade price as percent of notional'] = trade_price_percent_notional
    df['annual yield at strike at time of trade'] = annual_yield_at_strike
    df['yield at current mkt price at time of trade'] = yield_at_current_mkt_price_at_trade
    df['premium'] = premium
    df[f'contracted in {previous_5_months[4]}'] = month_5
    df[f'contracted in {previous_5_months[3]}'] = month_1
    df[f'contracted in {previous_5_months[2]}'] = month_2
    df[f'contracted in {previous_5_months[1]}'] = month_3
    df[f'contracted in {previous_5_months[0]}'] = month_4
    df['trade date'] = trade_date
    df['days till exp (trade date)'] = days_till_exp_date
    df['days till exp (current)'] = days_till_exp_date_current
    df['underlying price at time of trade'] = underlying_price_at_time_of_trade
    df['otm at time of trade'] = otm_at_time_of_trade
    df['underlying price, current'] = underlying_price_current
    df['otm, current'] = otm_current
    df['$ amount of stock itm can be called (-) or put (+)'] = amount_of_stock_itm_can_be_called
    df['weight'] = weight
    df['weighted otm'] = weighted_otm
    df['mkt beta'] = mkt_beta_list
    df['cash if exercised'] = cash_if_exercised
    df['=AK1-A1'] = week_1
    df['=AL1-A1'] = week_2
    df['=AM1-A1'] = week_3
    df['=AN1-A1'] = week_4
    df['=AO1-A1'] = week_5
    df['=AP1-A1'] = week_6
    df['=AQ1-A1'] = week_7
    df['=AR1-A1'] = week_8
    df['=AS1-A1'] = week_9
    df['=AT1-A1'] = week_10
    df['=AU1-A1'] = week_12
    df.insert(0, 'check date >>', '')  # or use an empty string: ''
'''

# Extract content between single quotes and concatenate into a list
values = re.findall(r"'([^']*)'", data)

print(values)

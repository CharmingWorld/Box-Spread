import pandas as pd
import numpy as np
import itertools
import math

data = pd.read_excel(r'沪深300股指期权盒式套利.xlsx')
shibor_data = pd.read_excel(r'Shibor历史数据.xlsx')
date = []  # 日期
due_date = []  # 到期日

# 生成日期列表
for i in range(len(data['日期'])):
    date.append(str(data['日期'][i])[0:10])
    date = sorted(set(date))  # 日期列表

# 生成到期日列表
for i in range(len(data['到期日'])):
    due_date.append(data['到期日'][i])
    due_date = sorted(set(due_date))  # 到期日列表


#  生成当天所有可能的盒式期权组合(相同日期、相同到期日、不同行权价的期权组合)
def epc(date, due_date, data):
    print('正在生成 {} 的所有到期日为 {} 的套利组合\n'.format(date, due_date))
    all_exercise_price_options = []
    portfolio = []
    # 生成当天所有行权价的期权列表
    for i in range(len(data['日期'])):
        if date == str(data['日期'][i])[0:10] and due_date == data['到期日'][i]:
            per_exercise_price_options = [data['行权价'][i], data['C收盘价'][i], data['P收盘价'][i]]
            all_exercise_price_options.append(per_exercise_price_options)
    # 根据当天不同行权价进行期权组合
    for i in itertools.combinations(all_exercise_price_options, 2):
        portfolio.append(i)
    portfolio = np.array(portfolio)
    print('本次套利组合生成完毕\n')
    return portfolio


# 进行当天套利操作
def arbitrage(portfolio, r, t, all_margin_used, all_profit, per_date, per_due_date):
    if len(portfolio) != 0:
        for i in range(len(portfolio)):
            print('正在进行 {} 日的到期日为 {} 的组合的第 {} 次套利\n'.format(per_date, per_due_date, i + 1))
            value_X2_X1 = 100 * (portfolio[i][1, 0] - portfolio[i][0, 0]) * (math.e ** (-1 * r * t / 360))  # X2-X1
            value_C_P = 100 * (portfolio[i][0, 1] - portfolio[i][1, 1] + portfolio[i][1, 2] - portfolio[i][
                0, 2])  # -100*(C1-C2=P2-P1)
            value_portfolio_long = value_X2_X1 - value_C_P - 4 * 15 - 2 * 2 * (
                    math.e ** (-1 * r * t / 360))  # 多头盒式价差减去期初手续费和期末行权费现值
            value_portfolio_short = value_C_P - value_X2_X1 - 4 * 15 - 2 * 2 * (
                    math.e ** (-1 * r * t / 360))  # 空头盒式价差减去期初手续费和期末行权费现值

            if value_portfolio_long > 0:  # 进行多头盒式套利
                margin_used = 0.12 * 100 * (portfolio[i, 0][1] + portfolio[i, 1][2])  # 做空期权交纳的保证金数量
                profit = value_portfolio_long * (math.e ** (r * t / 360))  # 多头盒式套利到期收益额
                all_margin_used.append(round(margin_used, 0))
                all_profit.append(round(profit, 0))

            elif value_portfolio_short > 0:  # 进行空头盒式套利
                margin_used = 0.12 * 100 * (portfolio[i, 0][2] + portfolio[i, 1][1])
                profit = value_portfolio_short * (math.e ** (r * t / 360))  # 空头盒式套利到期收益
                all_margin_used.append(round(margin_used, 0))
                all_profit.append(round(profit, 0))
            else:  # 不进行套利
                print('第 {} 日的到期日为 {} 的组合第 {} 次套利没有套利机会。\n'.
                      format(per_date, per_due_date, i + 1))
                continue
            print('第 {} 日的到期日为 {} 的组合第 {} 次套利完成，本次使用保证金 {} 元，套利额为 {} 元。\n'.
                  format(per_date, per_due_date, i + 1, round(margin_used, 0), round(profit, 0)))
    else:
        print('{} 日没有到期日为 {} 的期权。\n'.format(per_date, per_due_date))


# 获取今日shibor无风险利率
def judge_risk_free_interest_rate(per_date, t, shibor_data):
    day = {1: 'O/N', 7: '1W', 14: '2W', 30: '1M', 90: '3M', 180: '6M', 270: '9M', 360: '1Y'}
    day_used = (min(day.keys(), key=lambda x: abs(x - t)))
    r = shibor_data[shibor_data['日期'] == per_date][day[day_used]]
    return 0.01 * float(r)



day_margin_used = []
day_profit = []
for per_date in date:
    margin_used = []
    profit = []
    for per_due_date in due_date:
        portfolio = epc(per_date, per_due_date, data)  # 生成当天所有可能的盒式期权组合
        all_margin_used = []
        all_profit = []
        t = 30 * (int(str(per_due_date)[2:4]) - int(per_date[5:7])) + 15 - int(per_date[8:10])  # 期权到期时间
        r = judge_risk_free_interest_rate(per_date, t, shibor_data)  # 获取今日shibor无风险利率
        arbitrage(portfolio, r, t, all_margin_used, all_profit, per_date, per_due_date)
        margin_used.append(sum(all_margin_used))
        profit.append((sum(all_profit)))
    day_margin_used.append(sum(margin_used))
    day_profit.append(sum(profit))

principal = sum(day_margin_used)
profit = sum(day_profit)
# 计算收益率
print('盒式套利的收益率为：{:.2f}%'.format(100 * profit / principal))
# 输出套利结果
results_dic = {'日期': date, '每日使用保证金': day_margin_used, '每日套利收益': day_profit}
arbitrage_results = pd.DataFrame(results_dic)
arbitrage_results.set_index('日期')
print(arbitrage_results)
arbitrage_results.to_csv(r'套利结果.csv', index=True, header=True, sep=',', encoding='utf-8')

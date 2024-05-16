#判断本期是否营业
def is_open(days):
    if days > 0:
        return 1
    else:
        return 0

#数字转换为百分比
def convert_to_percentage(number):  
    return f"{number:.2%}"  

#数字转换为“万”
def num_to_wan(num):  
    if not isinstance(num, (int, float)):  
        print(num)
        raise ValueError("输入必须为数字")
    return f'{round(num / 10000, 2)}万'

#计算同环比
def Tonghuanbi(current_period,previous_period):
    
    return (current_period - previous_period)/ previous_period

#保留两位小数
def format_number(num):  
    num = round(num, 2)  
    return format(num, '.2f')

#增加&减少判断
def in_or_decrease(num):
    if str(num)[0]=='-':
        return f'减少{str(num)[1:]}'
    else:
        return f'增加{num}'
#全量df筛选      
def generate_tables(quanliang_df):
    benqi_df = quanliang_df[quanliang_df['营业天数_本期'] > 0] #本期营业表
    huanbi_df = quanliang_df[quanliang_df['营业天数_环比期'] > 0] #环比期营业表
    tongbi_df = quanliang_df[quanliang_df['营业天数_同比期'] > 0] #同比期营业表
    huan_cun_df = quanliang_df[(quanliang_df['营业天数_本期'] > 0) & (quanliang_df['营业天数_环比期'] > 0)] #环比存量店铺
    tong_cun_df = quanliang_df[(quanliang_df['营业天数_本期'] > 0) & (quanliang_df['营业天数_同比期'] > 0) & (quanliang_df['运营状态'] == '营业中')] #同比存量店铺
    tonghuan_cun_df = quanliang_df[(quanliang_df['营业天数_本期'] > 0) & (quanliang_df['营业天数_环比期'] > 0)&(quanliang_df['营业天数_同比期'] > 0)] #同比存量店铺
    return quanliang_df,benqi_df,huanbi_df,tongbi_df,huan_cun_df,tong_cun_df,tonghuan_cun_df

#负数加括号
def fushu_kuohao(num):
    if num[0]=='-':
        return f'({num[1:]})'
    else:
        return num

#占比百分之多少
def zhanbi(x,y):
    return convert_to_percentage(x/y)


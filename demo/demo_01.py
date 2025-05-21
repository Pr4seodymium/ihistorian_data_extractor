import win32com.client

# 输入iHistorian数据库的相关信息
# iHistorian数据库使用windows身份验证，用户名和密码为windows的用户名和密码
# data_source为iHistorian数据库的名称，user_name为用户名，password为密码
data_source = "iHistorian Server"
user_name = "administrator"
password = "password"
 
# 连接iHistorian数据库
conn = win32com.client.Dispatch(r'ADODB.Connection')
# Provider工具为iHistorian客户端提供的"iHistorian OLE DB Provider"(非常重要)
conn.Open(f'PROVIDER=iHistorian OLE DB Provider;DATA SOURCE={data_source};USER ID={user_name};PASSWORD={password}')

# 读取iHistorian数据库中的数据
rs = win32com.client.Dispatch(r'ADODB.Recordset')
# 使用SQL语句查询iHistorian数据库中的数据
rs.Open("SELECT scada.steam_FT.F_CV.Value, timestamp FROM ihTrend WHERE timestamp > \"2024-11-01 00:00:00\" AND timestamp < \"2024-11-10 00:00:00\"",conn)

# 输出iHistorian数据库中的数据
# 具体使用方式参考可以pywin32的文档和ADO(ActiveX Data Objects)的文档
while not rs.EOF:
    for field in rs.Fields:
        print(f"{field.Name}: {field.Value}")
        i = i + 1
    rs.MoveNext()

# 关闭连接
rs.Close()
conn.Close()

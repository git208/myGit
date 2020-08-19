import TestScriptNew

path_down = '/Users/far/Downloads/bbbbkkkk'
path_monitor = '/Users/far/Desktop/bankuaicehsi'
result_name = '成分股测试.xlsx'
bankuai_name = {
    'Trade_fz':'',
    'Notion_obvious':'BanKuai_',
    'Trade_sw_obvious':'',
    'Trade_sw1':'',
    'Area_szyp':'',
    'Notion_szyp':''
}



TestScriptNew.excle_more(path_down,result_name,bankuai_name,True)
print('*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-\n\n\n\n*-*-*-*--*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-')
TestScriptNew.excle_monitor(path_monitor,bankuai_name,result_name)

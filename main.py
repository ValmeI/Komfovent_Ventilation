from functions import get_vent_stats, add_xpos_in_list, column_width, create_excel, write_to_excel
import datetime

'# get today-s datetime but without milliseconds'
today_str = datetime.datetime.now().isoformat(' ', 'seconds')

'# add data from api-s to one big list. You need to know your ventilation local IP aadress'
combined_data_list = get_vent_stats(komfovent_local_ip="http://192.168.50.78/", var='det') + \
                     get_vent_stats(komfovent_local_ip="http://192.168.50.78/", var='i2') + \
                     get_vent_stats(komfovent_local_ip="http://192.168.50.78/", var='i')

'# add today-s date the beginning of the list'
new_vent_data_list = add_xpos_in_list(var=today_str, pos=0, input_list=combined_data_list)
write_to_excel(excel_name="SampleData", list_of_data=new_vent_data_list)
print(f'Success: {today_str}')
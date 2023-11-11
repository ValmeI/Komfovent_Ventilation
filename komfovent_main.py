import datetime

from functions import get_vent_stats, add_xpos_in_list, write_to_excel

if __name__ == '__main__':
    print('Starting...')
    IP = "http://192.168.50.78/"

    # Add data from APIs to one big list. You need to know your ventilation local IP address.
    combined_data_list = (
        get_vent_stats(komfovent_local_ip=IP, var='det') +
        get_vent_stats(komfovent_local_ip=IP, var='i2') +
        get_vent_stats(komfovent_local_ip=IP, var='i')
    )

    # Add today's date to the beginning of the list
    today_str = datetime.datetime.now().isoformat(' ', 'seconds')
    new_vent_data_list = add_xpos_in_list(var=today_str, pos=0, input_list=combined_data_list)
    write_to_excel(excel_name="SampleData", list_of_data=new_vent_data_list)
    print(f'Success: {today_str}')

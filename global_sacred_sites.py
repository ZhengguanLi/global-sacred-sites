import re
import xlwt
import requests 

def get_latlng(url):
	""" get the latitude and longtitude result"""

	web_data = requests.get(url)
	pattern_str = r'var titlePlacemark\d+ = "(.*)";\nvar latlng\d+ = new google.maps.LatLng\((.*), (.*)\)'
	pattern = re.compile(pattern_str)
	res = re.findall(pattern, web_data.text) 
	
	print('Found '+ str(len(res)) + ' results')

	# remove leading and trailing whitespace characters
	for index in range(len(res)):
		temp = list(res[index])
		temp[0] = str.strip(temp[0])
		res[index] = tuple(temp)
	
	# sort the data in ascending order(alphabet)
	list.sort(res, key = lambda a_tuple: a_tuple[0])
	
	return res
	
def export(res):
	""" export data to excel file """

	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet('sheet 1')
	
	# set font
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = 'Times New Roman'
	font.height = 20 * 12
	style.font = font
	
	# set the width of the colomn
	for col_index in range(3):
		sheet.col(col_index).width = 256 * 30
	
	# set the height of row
	for row_index in range(len(res)):
		sheet.row(row_index).height_mismatch = True
		sheet.row(row_index).height = 350
	
	print('Exporting to excel...')

	# write header to excel
	sheet.write(0, 0, 'City', style)
	sheet.write(0, 1, 'Latitude', style)
	sheet.write(0, 2, 'Longtitude', style)
	
	# write data(city_name, latitude, longtitude) to excel
	i = 1
	for per_info in res:
		sheet.write(i, 0, per_info[0], style) # city name
		sheet.write(i, 1, per_info[1], style) # latitude
		sheet.write(i, 2, per_info[2], style) # longtitude
		i = i + 1

	workbook.save('global_sacred_sites.xls')
	print('Export done!')

if __name__ == '__main__':
	url = "https://sacredsites.com/global_sacred_sites.html"
	res = get_latlng(url)	
	export(res)
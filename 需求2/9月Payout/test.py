import json
import re

if __name__ == '__main__':
	# with open('退款筛选条件.txt', 'r', encoding='utf-8') as file:
	# 	for line in file:
	# 		t_dict = json.loads(line)
	# 		if t_dict.get('c_cell'):
	# 			error_dict = t_dict.get('c_cell')
	# 		else:
	# 			lack_dict = t_dict.get('ak_cell')
	# print(error_dict)
	# print(lack_dict)

	abc = [[[[26.0, 42.0], [117.0, 42.0], [117.0, 65.0], [26.0, 65.0]], ('Payout date', 0.9731245040893555)], [[[275.0, 42.0], [350.0, 42.0], [350.0, 65.0], [275.0, 65.0]], (' Payout ID', 0.9756097793579102)], [[[456.0, 43.0], [566.0, 43.0], [566.0, 62.0], [456.0, 62.0]], ('Payout method', 0.9969693422317505)], [[[749.0, 42.0], [800.0, 42.0], [800.0, 65.0], [749.0, 65.0]], ('Status', 0.9996630549430847)], [[[953.0, 38.0], [1005.0, 42.0], [1003.0, 67.0], [951.0, 63.0]], ('Memo', 0.9990884065628052)], [[[1414.0, 42.0], [1478.0, 42.0], [1478.0, 65.0], [1414.0, 65.0]], ('Amount', 0.999946117401123)], [[[24.0, 92.0], [184.0, 90.0], [184.0, 113.0], [24.0, 115.0]], ('Sep 30, 2023 16:06:43', 0.9869534373283386)], [[[275.0, 93.0], [363.0, 93.0], [363.0, 111.0], [275.0, 111.0]], ('6052841673', 0.9998943209648132)], [[[453.0, 92.0], [613.0, 90.0], [613.0, 113.0], [453.0, 115.0]], (' Payoneer ID 44113246', 0.9800914525985718)], [[[748.0, 88.0], [847.0, 92.0], [846.0, 115.0], [747.0, 111.0]], (' Funds sent', 0.9992191791534424)], [[[952.0, 90.0], [1189.0, 92.0], [1189.0, 115.0], [952.0, 113.0]], (' Funds usually arrive within the day', 0.9670202732086182)], [[[1426.0, 91.0], [1478.0, 91.0], [1478.0, 115.0], [1426.0, 115.0]], ("66'69$", 0.9741852283477783)], [[[27.0, 141.0], [181.0, 141.0], [181.0, 163.0], [27.0, 163.0]], ('Sep 29, 2023 16:11:03', 0.9998124241828918)], [[[278.0, 143.0], [365.0, 143.0], [365.0, 161.0], [278.0, 161.0]], ('6052943313', 0.9999052882194519)], [[[454.0, 141.0], [613.0, 141.0], [613.0, 163.0], [454.0, 163.0]], ('Payoneer ID 44113246', 0.9725039601325989)], [[[747.0, 141.0], [846.0, 141.0], [846.0, 165.0], [747.0, 165.0]], ('0 Funds sent', 0.9439485669136047)], [[[955.0, 141.0], [1190.0, 141.0], [1190.0, 163.0], [955.0, 163.0]], ('Funds usually arrive within the day', 0.977765679359436)], [[[1418.0, 140.0], [1478.0, 140.0], [1478.0, 165.0], [1418.0, 165.0]], ('$246.98', 0.9999659657478333)], [[[26.0, 190.0], [182.0, 190.0], [182.0, 211.0], [26.0, 211.0]], ('Sep 28, 2023 16:09:17', 0.9979813098907471)], [[[275.0, 191.0], [363.0, 191.0], [363.0, 210.0], [275.0, 210.0]], (' 6051729033', 0.9685933589935303)], [[[454.0, 190.0], [613.0, 190.0], [613.0, 211.0], [454.0, 211.0]], ('Payoneer ID 44113246', 0.9987931251525879)], [[[749.0, 191.0], [845.0, 191.0], [845.0, 210.0], [749.0, 210.0]], ('o Funds sent', 0.9475316405296326)], [[[954.0, 190.0], [1190.0, 190.0], [1190.0, 211.0], [954.0, 211.0]], (' Funds usually arrive within the day', 0.9859732389450073)], [[[1418.0, 190.0], [1478.0, 190.0], [1478.0, 213.0], [1418.0, 213.0]], ('$168.38', 0.9999027848243713)], [[[27.0, 240.0], [181.0, 240.0], [181.0, 261.0], [27.0, 261.0]], ('Sep 27, 2023 16:07:11', 0.9924233555793762)], [[[277.0, 241.0], [365.0, 241.0], [365.0, 259.0], [277.0, 259.0]], ('6050576625', 0.9998027086257935)], [[[454.0, 240.0], [613.0, 240.0], [613.0, 261.0], [454.0, 261.0]], ('Payoneer ID 44113246', 0.9988336563110352)], [[[749.0, 241.0], [845.0, 241.0], [845.0, 259.0], [749.0, 259.0]], ('0 Funds sent', 0.9465880393981934)], [[[955.0, 240.0], [1190.0, 240.0], [1190.0, 261.0], [955.0, 261.0]], ('Funds usually arrive within the day', 0.9889543652534485)], [[[1418.0, 240.0], [1478.0, 240.0], [1478.0, 263.0], [1418.0, 263.0]], ('$222.08', 0.9998583793640137)], [[[26.0, 288.0], [184.0, 288.0], [184.0, 309.0], [26.0, 309.0]], ('Sep 26, 2023 16:05:44', 0.993414044380188)], [[[277.0, 289.0], [363.0, 289.0], [363.0, 308.0], [277.0, 308.0]], ('6049069881', 0.9996681213378906)], [[[454.0, 288.0], [613.0, 288.0], [613.0, 309.0], [454.0, 309.0]], ('Payoneer ID 44113246', 0.9987931251525879)], [[[748.0, 284.0], [847.0, 288.0], [846.0, 311.0], [747.0, 307.0]], (' Funds sent', 0.9992191791534424)], [[[954.0, 289.0], [1190.0, 289.0], [1190.0, 311.0], [954.0, 311.0]], ('Funds usually arrive within the day', 0.9877769351005554)], [[[1426.0, 288.0], [1478.0, 288.0], [1478.0, 311.0], [1426.0, 311.0]], ('$98.08', 0.999962568283081)], [[[27.0, 338.0], [184.0, 338.0], [184.0, 359.0], [27.0, 359.0]], ('Sep 25, 2023 16:06:28', 0.9977679252624512)], [[[275.0, 336.0], [366.0, 336.0], [366.0, 359.0], [275.0, 359.0]], ('6046479849', 0.999758243560791)], [[[454.0, 338.0], [613.0, 338.0], [613.0, 359.0], [454.0, 359.0]], ('Payoneer ID 44113246', 0.9993880391120911)], [[[746.0, 334.0], [847.0, 338.0], [846.0, 363.0], [745.0, 359.0]], (' Funds sent', 0.9989063739776611)], [[[955.0, 338.0], [1190.0, 338.0], [1190.0, 359.0], [955.0, 359.0]], ('Funds usually arrive within the day', 0.9889543652534485)], [[[1418.0, 336.0], [1478.0, 336.0], [1478.0, 361.0], [1418.0, 361.0]], ('$224.42', 0.999968945980072)], [[[26.0, 386.0], [182.0, 386.0], [182.0, 407.0], [26.0, 407.0]], ('Sep 24, 2023 16:08:41', 0.9977092742919922)], [[[275.0, 388.0], [362.0, 388.0], [362.0, 406.0], [275.0, 406.0]], ('6046472841', 0.9628831744194031)], [[[454.0, 386.0], [613.0, 386.0], [613.0, 407.0], [454.0, 407.0]], ('Payoneer ID 44113246', 0.9987931251525879)], [[[749.0, 388.0], [845.0, 388.0], [845.0, 406.0], [749.0, 406.0]], ('o Funds sent', 0.9481041431427002)], [[[955.0, 386.0], [1190.0, 388.0], [1190.0, 409.0], [955.0, 407.0]], ('Funds usually arrive within the day', 0.9997950792312622)], [[[1418.0, 386.0], [1478.0, 386.0], [1478.0, 409.0], [1418.0, 409.0]], ('$172.39', 0.9999799132347107)], [[[26.0, 436.0], [182.0, 436.0], [182.0, 457.0], [26.0, 457.0]], ('Sep 23, 2023 16:05:01', 0.998671293258667)], [[[277.0, 437.0], [363.0, 437.0], [363.0, 456.0], [277.0, 456.0]], ('6045194721', 0.9999722242355347)], [[[454.0, 436.0], [613.0, 436.0], [613.0, 457.0], [454.0, 457.0]], ('Payoneer ID 44113246', 0.9988336563110352)], [[[750.0, 437.0], [845.0, 437.0], [845.0, 456.0], [750.0, 456.0]], ('o Funds sent', 0.9586388468742371)], [[[955.0, 436.0], [1190.0, 436.0], [1190.0, 457.0], [955.0, 457.0]], ('Funds usually arrive within the day', 0.9889543652534485)], [[[1311.0, 431.0], [1351.0, 435.0], [1348.0, 470.0], [1308.0, 466.0]], ('de', 0.9991652369499207)], [[[1417.0, 430.0], [1479.0, 435.0], [1477.0, 460.0], [1415.0, 455.0]], ('$110.93', 0.9999524354934692)]]
	dict_1 = {}
	res = []
	current_year = '2023'
	month_dict = {
    '1': 'Jan', '2': 'Feb', '3': 'Mar', '4': 'Apr',
    '5': 'May', '6': 'Jun', '7': 'Jul', '8': 'Aug',
    '9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
	}
	check_month_list = [month_dict[str(9)], month_dict[str(8)], month_dict[str(10)]]


	pattern = r"\$\d+"

	for sub_item in abc:
		# if '$' in sub_item[1][0]:
		if re.match(r'US\s*\$|US\s*S|USS|US\$', sub_item[1][0]):
			res_1 = sub_item[1][0].replace(' ', '')
			res.append(re.split(r'USS|US\$', res_1)[1])
		elif re.search(r"\$\d+.*", sub_item[1][0]):
			res.append(re.search(r"\$\d+.*", sub_item[1][0])[0].split('$')[1])
		elif "'" in sub_item[1][0] and "$" in sub_item[1][0]:
			res.append("识别有误，请手动填写")
		elif 'Start' not in sub_item[1][0] and 'End' not in sub_item[1][0]:
			res_1 = sub_item[1][0].replace(' ', '').replace(',', '').replace('.', '')
			for cml in check_month_list:
				if cml in res_1:
					res.append(re.split(r'' + str(current_year) + '', res_1)[0])
		#
		# elif '.' + current_year in sub_item[1][0]:
		# 	if ' ' in sub_item[1][0].split('.')[0]:
		# 		# print(123)
		# 		res.append(sub_item[1][0].split('.')[0].split()[1])
		# 	else:
		# 		res.append(sub_item[1][0].split('.')[0].split(month_dict[month_content])[1])
		# elif month_dict[month_content] in sub_item[1][0] and 'Start' not in sub_item[1][0] and 'End' not in sub_item[1][
		# 	0]:
		# 	if ' ' in sub_item[1][0].split(',')[0]:
		# 		if current_year in sub_item[1][0].split(',')[0].split()[0]:
		# 			res.append(
		# 				sub_item[1][0].split(',')[0].split()[0].split(current_year)[0].split(month_dict[month_content])[
		# 					1])
		# 		elif current_year in sub_item[1][0].split(',')[0].split()[1]:
		# 			res.append(sub_item[1][0].split(',')[0].split()[1].split(current_year)[0])
		# 		else:
		# 			res.append(sub_item[1][0].split(',')[0].split()[1])
		# 	else:
		# 		res.append(sub_item[1][0].split(',')[0].split(month_dict[month_content])[1])
	for i in range(0, len(res), 2):
		key = res[i]
		value = res[i + 1]
		dictionary = {key: value}
		dict_1.update(dictionary)

	filtered_dict = {key.split(month_dict[str(9)])[1]: value for key, value in dict_1.items() if month_dict[str(9)] in key}

	print(filtered_dict)
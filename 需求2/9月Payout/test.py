data = [
	[
		[
			[
				[556.0, 3.0],
				[575.0, 3.0],
				[575.0, 20.0],
				[556.0, 20.0]
			], ('AA', 0.9951339364051819)
		],
		[
			[
				[649.0, 2.0],
				[707.0, 2.0],
				[707.0, 20.0],
				[649.0, 20.0]
			], ('eBay-', 0.9261108636856079)
		],
		[
			[
				[782.0, 0.0],
				[921.0, 0.0],
				[921.0, 22.0],
				[782.0, 22.0]
			], ('47.119. 4. 6 ( )', 0.9072273969650269)
		],
		[
			[
				[160.0, 29.0],
				[304.0, 29.0],
				[304.0, 50.0],
				[160.0, 50.0]
			], ('*', 0.5959026217460632)
		],
		[
			[
				[314.0, 30.0],
				[457.0, 30.0],
				[457.0, 47.0],
				[314.0, 47.0]
			], ('M Overview - eBay S..', 0.8819522857666016)
		],
		[
			[
				[469.0, 32.0],
				[610.0, 32.0],
				[610.0, 49.0],
				[469.0, 49.0]
			], (' Manage active lis*-', 0.9156426191329956)
		],
		[
			[
				[622.0, 29.0],
				[767.0, 29.0],
				[767.0, 50.0],
				[622.0, 50.0]
			], (' Seller Hub | Your', 0.963664174079895)
		],
		[
			[
				[779.0, 32.0],
				[879.0, 32.0],
				[879.0, 49.0],
				[779.0, 49.0]
			], ('FAR', 0.6873882412910461)
		],
		[
			[
				[889.0, 32.0],
				[1030.0, 32.0],
				[1030.0, 49.0],
				[889.0, 49.0]
			], (' Block buyers fron-', 0.9182206392288208)
		],
		[
			[
				[1046.0, 32.0],
				[1144.0, 32.0],
				[1144.0, 49.0],
				[1046.0, 49.0]
			], ('*', 0.5360903143882751)
		],
		[
			[
				[1154.0, 32.0],
				[1284.0, 32.0],
				[1284.0, 49.0],
				[1154.0, 49.0]
			], (' Manage pronotions', 0.9925505518913269)
		],
		[
			[
				[36.0, 59.0],
				[158.0, 59.0],
				[158.0, 76.0],
				[36.0, 76.0]
			], ('Sep 19, 2023 16:12:21', 0.999729335308075)
		],
		[
			[
				[233.0, 57.0],
				[306.0, 57.0],
				[306.0, 76.0],
				[233.0, 76.0]
			], ('6037461656', 0.9993680715560913)
		],
		[
			[
				[462.0, 55.0],
				[605.0, 55.0],
				[605.0, 77.0],
				[462.0, 77.0]
			], ('Payoneer KS 42988423', 0.9369204640388489)
		],
		[
			[
				[1518.0, 57.0],
				[1579.0, 57.0],
				[1579.0, 76.0],
				[1518.0, 76.0]
			], ('US $43.07', 0.9887667894363403)
		],
]]

result = []
month = ['Sep','US $']


for sub_item in data[0]:
	for item in month:
		if item in sub_item[1][0]:
			result.append(sub_item[1][0])
print(result)

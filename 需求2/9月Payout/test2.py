ABC = [
'"MO","64108-1105","US","10.06","USD","Sep 17, 2023","6034740334","East West Bank *8924","Funds sent","--","314662346419","1341795287021","26'' 26x1.90/26x1.95/26x2/26x2.10/26x2.125 Inch Bicycle Inner Tube MTB Bike Tire","esuk|A1362#53","1","11.95","0","--","0.89","-0.3","-1.59","--","--","--","--","11.95","USD","--","--","--"',
'"CA","93305-5621","US","2.26","USD","Sep 17, 2023","6034740334","East West Bank *8924","Funds sent","--","314091994704","1341795585021","2-PACK For iPhone 15 14 Pro Max 13 12 11 XS Max Tempered GLASS Screen Protector","E04552","1","2.96","0","--","0.24","-0.3","-0.4","--","--","--","--","2.96","USD","--","--","--"',
'"CA","94086-7401","US","2.25","USD","Sep 17, 2023","6034740334","East West Bank *8924","Funds sent","--","314821774245","1341795209021","2X Tempered Glass Screen Protector For iPhone 15 14 13 12 11 Pro Max X XS Max XR","esuk|E04547#81","1","2.95","0","--","0.27","-0.3","-0.4","--","--","--","--","2.95","USD","--","--","--"',
'"LUKEVILLE","AZ","85341-0453","US","2.26","USD","Sep 17, 2023","6034740334","East West Bank *8924","Funds sent","--","314821774245","1341795054021","2X Tempered Glass Screen Protector For iPhone 15 14 13 12 11 Pro Max X XS Max XR","esuk|E03936#70","1","2.95","0","--","0.18","-0.3","-0.39","--","--","--","--","2.95","USD","--","--","--"'
]


for line in ABC:
    if '","' in line[0]:
        new_line = line.replace('","', '@').replace('"', '').replace('\n', '')
        new_line = new_line.split('@')

        print([new_line])
    else:
        new_line = line.replace('"', '').replace('\n', '')

        print([new_line])
fullNames = [
'Aatrox',
'Ahri',
'Akali',
'Alistar',
'Amumu',
'Anivia',
'Annie',
'Ashe',
'Azir',
'Bard',
'Blitzcrank',
'Brand',
'Braum',
'Caitlyn',
'Cassiopeia',
'Chogath',
'Corki',
'Darius',
'Diana',
'Draven',
'Drmundo',
'Ekko',
'Elise',
'Evelynn',
'Ezreal',
'Fiddlesticks',
'Fiora',
'Fizz',
'Galio',
'Gangplank',
'Garen',
'Gnar',
'Gragas',
'Graves',
'Hecarim',
'Heimerdinger',
'Illaoi',
'Irelia',
'Janna',
'Jarvaniv',
'Jax',
'Jayce',
'Jinx',
'Kalista',
'Karma',
'Karthus',
'Kassadin',
'Katarina',
'Kayle',
'Kennen',
'Khazix',
'Kindred',
'Kogmaw',
'Leblanc',
'Leesin',
'Leona',
'Lissandra',
'Lucian',
'Lulu',
'Lux',
'Malphite',
'Malzahar',
'Maokai',
'Masteryi',
'Missfortune',
'Mordekaiser',
'Morgana',
'Nami',
'Nasus',
'Nautilus',
'Nidalee',
'Nocturne',
'Nunu',
'Olaf',
'Orianna',
'Pantheon',
'Poppy',
'Quinn',
'Rammus',
'Reksai',
'Renekton',
'Rengar',
'Riven',
'Rumble',
'Ryze',
'Sejuani',
'Shaco',
'Shen',
'Shyvana',
'Singed',
'Sion',
'Sivir',
'Skarner',
'Sona',
'Soraka',
'Swain',
'Syndra',
'Tahmkench',
'Talon',
'Taric',
'Teemo',
'Thresh',
'Tristana',
'Trundle',
'Tryndamere',
'Twistedfate',
'Twitch',
'Udyr',
'Urgot',
'Varus',
'Vayne',
'Veigar',
'Velkoz',
'Vi',
'Viktor',
'Vladimir',
'Volibear',
'Warwick',
'Wukong',
'Xerath',
'Xinzhao',
'Yasuo',
'Yorick',
'Zac',
'Zed',
'Ziggs',
'Zilean',
'Zyra'
]
shortenedNames = [
'Aatr',
'Ahri',
'Akali',
'Alis',
'Amum',
'Aniv',
'Annie',
'Ashe',
'Azir',
'Bard',
'Blitz',
'Brand',
'Braum',
'Cait',
'Cass',
'Chog',
'Cork',
'Dari',
'Diana',
'Drav',
'Mundo',
'Ekko',
'Elise',
'Evel',
'Ezre',
'Fidd',
'Fior',
'Fizz',
'Gali',
'Gp',
'Garen',
'Gnar',
'Grag',
'Grav',
'Hec',
'Heim',
'Illa',
'Irel',
'Jann',
'Jarv',
'Jax',
'Jayce',
'Jinx',
'Kali',
'Karma',
'Karth',
'Kass',
'Kata',
'Kayle',
'Kenn',
'Kha',
'Kind',
'Kog',
'Lb',
'Lee',
'Leo',
'Liss',
'Luc',
'Lulu',
'Lux',
'Malp',
'Malz',
'Mao',
'Yi',
'Mf',
'Morde',
'Morg',
'Nami',
'Nas',
'Naut',
'Nid',
'Noct',
'Nunu',
'Olaf',
'Ori',
'Panth',
'Popp',
'Quinn',
'Ram',
'Rek',
'Renek',
'Ren',
'Riven',
'Rum',
'Ryze',
'Sej',
'Shaco',
'Shen',
'Shyv',
'Sing',
'Sion',
'Siv',
'Skar',
'Sona',
'Raka',
'Swain',
'Syn',
'Tahm',
'Talon',
'Taric',
'Teemo',
'Thre',
'Trist',
'Trun',
'Trynd',
'Tf',
'Twit',
'Udyr',
'Urgot',
'Varus',
'Vayne',
'Veig',
'Vel',
'Vi',
'Vikt',
'Vlad',
'Voli',
'Warw',
'Wuko',
'Xera',
'Xinz',
'Yasu',
'Yori',
'Zac',
'Zed',
'Zigg',
'Zil',
'Zyra']

nameMap = {
    'Aatrox' : 'Aatr',
    'Ahri' : 'Ahri',
    'Akali' : 'Akali',
    'Alistar' : 'Ali',
    'Amumu' : 'Amum',
    'Anivia' : 'Aniv',
    'Annie' : 'Anni',
    'Ashe' : 'Ashe',
    'Azir' : 'Azir',
    'Bard' : 'Bard',
    'Blitzcrank' : 'Blitz',
    'Brand' : 'Bran',
    'Braum' : 'Brau',
    'Caitlyn' : 'Cait',
    'Cassiopeia' : 'Cass',
    'Chogath' : 'Chog',
    'Corki' : 'Corki',
    'Darius' : 'Dar',
    'Diana' : 'Diana',
    'Draven' : 'Drav',
    'Drmundo' : 'Mund',
    'Ekko' : 'Ekko',
    'Elise' : 'Eli',
    'Evelynn' : 'Eve',
    'Ezreal' : 'Ez',
    'Fiddlesticks' : 'Fidd',
    'Fiora' : 'Fio',
    'Fizz' : 'Fizz',
    'Galio' : 'Gali',
    'Gangplank' : 'GP',
    'Garen' : 'Garen',
    'Gnar' : 'Gnar',
    'Gragas' : 'Grag',
    'Graves' : 'Grav',
    'Hecarim' : 'Hec',
    'Heimerdinger' : 'Heim',
    'Illaoi' : 'Illa',
    'Irelia' : 'Ire',
    'Janna' : 'Janna',
    'Jarvaniv' : 'J4',
    'Jax' : 'Jax',
    'Jayce' : 'Jay',
    'Jinx' : 'Jinx',
    'Kalista' : 'Kali',
    'Karma' : 'Karma',
    'Karthus' : 'Karth',
    'Kassadin' : 'Kass',
    'Katarina' : 'Kata',
    'Kayle' : 'Kayle',
    'Kennen' : 'Ken',
    'Khazix' : 'Kha',
    'Kindred' : 'Kind',
    'Kogmaw' : 'Kog',
    'Leblanc' : 'LB',
    'Leesin' : 'Lee',
    'Leona' : 'Leo',
    'Lissandra' : 'Liss',
    'Lucian' : 'Luc',
    'Lulu' : 'Lulu',
    'Lux' : 'Lux',
    'Malphite' : 'Malp',
    'Malzahar' : 'Malz',
    'Maokai' : 'Mao',
    'Masteryi' : 'Yi',
    'Missfortune' : 'MF',
    'Mordekaiser' : 'Morde',
    'Morgana' : 'Morg',
    'Nami' : 'Nami',
    'Nasus' : 'Nasus',
    'Nautilus' : 'Naut',
    'Nidalee' : 'Nid',
    'None' : 'None',
    'Nocturne' : 'Noc',
    'Nunu' : 'Nunu',
    'Olaf' : 'Olaf',
    'Orianna' : 'Ori',
    'Pantheon' : 'Panth',
    'Poppy' : 'Pop',
    'Quinn' : 'Quin',
    'Rammus' : 'Ram',
    'Reksai' : 'Rek',
    'Renekton' : 'Renek',
    'Rengar' : 'Rengar',
    'Riven' : 'Riven',
    'Rumble' : 'Rumb',
    'Ryze' : 'Ryze',
    'Sejuani' : 'Sej',
    'Shaco' : 'Shaco',
    'Shen' : 'Shen',
    'Shyvana' : 'Shyv',
    'Singed' : 'Sing',
    'Sion' : 'Sion',
    'Sivir' : 'Siv',
    'Skarner' : 'Skar',
    'Sona' : 'Sona',
    'Soraka' : 'Sor',
    'Swain' : 'Swain',
    'Syndra' : 'Syn',
    'Tahmkench' : 'Tahm',
    'Talon' : 'Talon',
    'Taric' : 'Taric',
    'Teemo' : 'Teemo',
    'Thresh' : 'Thre',
    'Tristana' : 'Trist',
    'Trundle' : 'Trun',
    'Tryndamere' : 'Trynd',
    'Twistedfate' : 'TF',
    'Twitch' : 'Twit',
    'Udyr' : 'Udyr',
    'Urgot' : 'Urgot',
    'Varus' : 'Varus',
    'Vayne' : 'Vayne',
    'Veigar' : 'Veig',
    'Velkoz' : 'Vel',
    'Vi' : 'Vi',
    'Viktor' : 'Vik',
    'Vladimir' : 'Vlad',
    'Vlad' : 'Vlad',
    'Volibear' : 'Voli',
    'Warwick' : 'WW',
    'Wukong' : 'Wuko',
    'Xerath' : 'Xera',
    'Xinzhao' : 'Xin',
    'Yasuo' : 'Yas',
    'Yorick' : 'Yori',
    'Zac' : 'Zac',
    'Zed' : 'Zed',
    'Ziggs' : 'Zigg',
    'Zilean' : 'Zil',
    'Zyra' : 'Zyra',
};

colorMap = {
'Aatrox' : '#B0171F',
'Ahri' : '#DC143C',
'Akali' : '#FFB6C1',
'Alistar' : '#FFAEB9',
'Amumu' : '#EEA2AD',
'Anivia' : '#CD8C95',
'Annie' : '#8B5F65',
'Ashe' : '#FFC0CB',
'Azir' : '#FFB5C5',
'Bard' : '#EEA9B8',
'Blitzcrank' : '#CD919E',
'Brand' : '#8B636C',
'Braum' : '#FF3E96',
'Caitlyn' : '#EE3A8C',
'Cassiopeia' : '#CD3278',
'Chogath' : '#8B2252',
'Corki' : '#FF69B4',
'Darius' : '#FF6EB4',
'Diana' : '#EE6AA7',
'Draven' : '#CD6090',
'Drmundo' : '#8B3A62',
'Ekko' : '#8B1C62',
'Elise' : '#C71585',
'Evelynn' : '#D02090',
'Ezreal' : '#DA70D6',
'Fiddlesticks' : '#FF83FA',
'Fiora' : '#CD69C9',
'Fizz' : '#8B4789',
'Galio' : '#D8BFD8',
'Gangplank' : '#EED2EE',
'Garen' : '#CDB5CD',
'Gnar' : '#8B7B8B',
'Gragas' : '#8A2BE2',
'Graves' : '#912CEE',
'Hecarim' : '#9370DB',
'Heimerdinger' : '#5D478B',
'Illaoi' : '#483D8B',
'Irelia' : '#4169E1',
'Janna' : '#436EEE',
'Jarvaniv' : '#CAE1FF',
'Jax' : '#1874CD',
'Jayce' : '#63B8FF',
'Jinx' : '#4F94CD',
'Kalista' : '#87CEFA',
'Karma' : '#A4D3EE',
'Karthus' : '#8DB6CD',
'Kassadin' : '#7EC0EE',
'Katarina' : '#008080',
'Kayle' : '#20B2AA',
'Kennen' : '#40E0D0',
'Khazix' : '#00C78C',
'Kindred' : '#00FA9A',
'Kogmaw' : '#00FF7F',
'Leblanc' : '#00CD66',
'Leesin' : '#54FF9F',
'Leona' : '#3D9140',
'Lissandra' : '#C1FFC1',
'Lucian' : '#7CCD7C',
'Lulu' : '#32CD32',
'Lux' : '#00CD00',
'Malphite' : '#008B00',
'Malzahar' : '#EEEE00',
'Maokai' : '#FFF68F',
'Masteryi' : '#EEE685',
'Missfortune' : '#EEE685',
'Mordekaiser' : '#FFEC8B',
'Morgana' : '#FFD700',
'Nami' : '#CDAD00',
'Nasus' : '#EEB422',
'Nautilus' : '#CD8500',
'Nidalee' : '#FFE7BA',
'Nocturne' : '#FFE4B5',
'Nunu' : '#CDC0B0',
'Olaf' : '#8B8378',
'Orianna' : '#ED9121',
'Pantheon' : '#CD6600',
'Poppy' : '#8B4500',
'Quinn' : '#EE9A49',
'Rammus' : '#CDAF95',
'Reksai' : '#CDC5BF',
'Renekton' : '#EE4000',
'Rengar' : '#CD3700',
'Riven' : '#8B2500',
'Rumble' : '#E9967A',
'Ryze' : '#FF8C69',
'Sejuani' : '#CD7054',
'Shaco' : '#FF7256',
'Shen' : '#FF6347',
'Shyvana' : '#CD4F39',
'Singed' : '#FFC1C1',
'Sion' : '#EEB4B4',
'Sivir' : '#CD9B9B',
'Skarner' : '#8B6969',
'Sona' : '#CD5C5C',
'Soraka' : '#FF6A6A ',
'Swain' : '#EE6363',
'Syndra' : '#8E388E',
'Tahmkench' : '#7171C6',
'Talon' : '#7D9EC0',
'Taric' : '#388E8E',
'Teemo' : '#8E8E38',
'Thresh' : '#C5C1AA',
'Tristana' : '#C67171',
'Trundle' : '#C1C1C1',
'Tryndamere' : '#00BFFF',
'Twistedfate' : '#00FF00',
'Twitch' : '#00EE00',
'Udyr' : '#00CD00',
'Urgot' : '#008B00',
'Varus' : '#008000',
'Vayne' : '#006400',
'Veigar' : '#308014',
'Velkoz' : '#7CFC00',
'Vi' : '#7FFF00',
'Viktor' : '#76EE00',
'Vladimir' : '#66CD00',
'Volibear' : '#458B00',
'Warwick' : '#ADFF2F',
'Wukong' : '#CAFF70',
'Xerath' : '#BCEE68',
'Xinzhao' : '#A2CD5A',
'Yasuo' : '#6E8B3D',
'Yorick' : '#556B2F',
'Zac' : '#6B8E23',
'Zed' : '#C0FF3E',
'Ziggs' : '#B3EE3A',
'Zilean' : '#9ACD32',
'Zyra' : '#698B22',
}



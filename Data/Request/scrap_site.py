import json
from pprint import pprint

import requests
from loguru import logger


def parse(art: str) -> dict:
    cookies = {
        'current_location_id': '3922',
        'current_city': '814',
        'current_location_data': 'a%3A4%3A%7Bs%3A5%3A%22chain%22%3Ba%3A2%3A%7Bi%3A0%3Bi%3A68%3Bi%3A1%3Bi%3A3922%3B%7Ds%3A4%3A%22name%22%3Bs%3A22%3A%22%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%22%3Bs%3A9%3A%22full_name%22%3Bs%3A57%3A%22%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%D0%B0%D1%8F%20%D0%BE%D0%B1%D0%BB%2C%20%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%22%3Bs%3A11%3A%22location_id%22%3Bi%3A3922%3B%7D',
        'BITRIX_SM_SALE_UID': '1492353103',
        '_userGUID': '0:l3vgspid:9vBNTj8aTUicjc1ITcyA3Q3K0eoi1ZS7',
        '__exponea_etc__': '3894f658-9300-e055-98d0-466b3376dfbc',
        '_dvs': '0:l50bu63t:DQ~kcPDRr_UjRuFeGIBehpjLw_wqN6EJ',
        'PHPSESSID': 'mf4r8hl1khtq4g9appialjt9q8',
        'BITRIX_SM_ab_test_multi': '%7B%22aa06%22%3A%7B%22ID%22%3A%227710124%22%2C%22NAME%22%3A%22aa06%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa07%22%3A%7B%22ID%22%3A%227710127%22%2C%22NAME%22%3A%22aa07%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa08%22%3A%7B%22ID%22%3A%227710128%22%2C%22NAME%22%3A%22aa08%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa09%22%3A%7B%22ID%22%3A%227710129%22%2C%22NAME%22%3A%22aa09%22%2C%22GROUP%22%3A%22A%22%7D%2C%22aa10%22%3A%7B%22ID%22%3A%227710131%22%2C%22NAME%22%3A%22aa10%22%2C%22GROUP%22%3A%22B%22%7D%2C%22KS%22%3A%7B%22ID%22%3A%228116774%22%2C%22NAME%22%3A%22KS%22%2C%22GROUP%22%3A%22%22%7D%2C%22reg_bk_and_yz%22%3A%7B%22ID%22%3A%228122416%22%2C%22NAME%22%3A%22reg_bk_and_yz%22%2C%22GROUP%22%3A%22B%22%7D%2C%223d%22%3A%7B%22ID%22%3A%228123834%22%2C%22NAME%22%3A%223d%22%2C%22GROUP%22%3A%22%22%7D%2C%22ar%22%3A%7B%22ID%22%3A%228123836%22%2C%22NAME%22%3A%22ar%22%2C%22GROUP%22%3A%22%22%7D%2C%22services%22%3A%7B%22ID%22%3A%228362535%22%2C%22NAME%22%3A%22services%22%2C%22GROUP%22%3A%22B%22%7D%2C%22rr_basket%22%3A%7B%22ID%22%3A%228362213%22%2C%22NAME%22%3A%22rr_basket%22%2C%22GROUP%22%3A%22%22%7D%2C%22kt_left%22%3A%7B%22ID%22%3A%228107330%22%2C%22NAME%22%3A%22kt_left%22%2C%22GROUP%22%3A%22B%22%7D%2C%22credit_new_widget%22%3A%7B%22ID%22%3A%228324699%22%2C%22NAME%22%3A%22credit_new_widget%22%2C%22GROUP%22%3A%22A%22%7D%2C%22cartpopap%22%3A%7B%22ID%22%3A%228270671%22%2C%22NAME%22%3A%22cartpopap%22%2C%22GROUP%22%3A%22B%22%7D%2C%22alter_sort%22%3A%7B%22ID%22%3A%228377079%22%2C%22NAME%22%3A%22alter_sort%22%2C%22GROUP%22%3A%22A%22%7D%2C%22dates%22%3A%7B%22ID%22%3A%227895663%22%2C%22NAME%22%3A%22dates%22%2C%22GROUP%22%3A%22%22%7D%2C%22anyquery%22%3A%7B%22ID%22%3A%228402103%22%2C%22NAME%22%3A%22anyquery%22%2C%22GROUP%22%3A%22A%22%7D%7D',
        'dSesn': '8ff576f8-67a9-186a-df1d-9e4efac4a33c',
        'iwaf_fingerprint': '55c49ce9f14c551ca02458bb6e3ca0b9',
    }

    headers = {
        'authority': 'hoff.ru',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'content-type': 'application/json;charset=UTF-8',
        # Requests sorts cookies= alphabetically
        # 'cookie': 'current_location_id=3922; current_city=814; current_location_data=a%3A4%3A%7Bs%3A5%3A%22chain%22%3Ba%3A2%3A%7Bi%3A0%3Bi%3A68%3Bi%3A1%3Bi%3A3922%3B%7Ds%3A4%3A%22name%22%3Bs%3A22%3A%22%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%22%3Bs%3A9%3A%22full_name%22%3Bs%3A57%3A%22%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%D0%B0%D1%8F%20%D0%BE%D0%B1%D0%BB%2C%20%D0%9D%D0%BE%D0%B2%D0%BE%D1%81%D0%B8%D0%B1%D0%B8%D1%80%D1%81%D0%BA%22%3Bs%3A11%3A%22location_id%22%3Bi%3A3922%3B%7D; BITRIX_SM_SALE_UID=1492353103; _userGUID=0:l3vgspid:9vBNTj8aTUicjc1ITcyA3Q3K0eoi1ZS7; __exponea_etc__=3894f658-9300-e055-98d0-466b3376dfbc; _dvs=0:l50bu63t:DQ~kcPDRr_UjRuFeGIBehpjLw_wqN6EJ; PHPSESSID=mf4r8hl1khtq4g9appialjt9q8; BITRIX_SM_ab_test_multi=%7B%22aa06%22%3A%7B%22ID%22%3A%227710124%22%2C%22NAME%22%3A%22aa06%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa07%22%3A%7B%22ID%22%3A%227710127%22%2C%22NAME%22%3A%22aa07%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa08%22%3A%7B%22ID%22%3A%227710128%22%2C%22NAME%22%3A%22aa08%22%2C%22GROUP%22%3A%22B%22%7D%2C%22aa09%22%3A%7B%22ID%22%3A%227710129%22%2C%22NAME%22%3A%22aa09%22%2C%22GROUP%22%3A%22A%22%7D%2C%22aa10%22%3A%7B%22ID%22%3A%227710131%22%2C%22NAME%22%3A%22aa10%22%2C%22GROUP%22%3A%22B%22%7D%2C%22KS%22%3A%7B%22ID%22%3A%228116774%22%2C%22NAME%22%3A%22KS%22%2C%22GROUP%22%3A%22%22%7D%2C%22reg_bk_and_yz%22%3A%7B%22ID%22%3A%228122416%22%2C%22NAME%22%3A%22reg_bk_and_yz%22%2C%22GROUP%22%3A%22B%22%7D%2C%223d%22%3A%7B%22ID%22%3A%228123834%22%2C%22NAME%22%3A%223d%22%2C%22GROUP%22%3A%22%22%7D%2C%22ar%22%3A%7B%22ID%22%3A%228123836%22%2C%22NAME%22%3A%22ar%22%2C%22GROUP%22%3A%22%22%7D%2C%22services%22%3A%7B%22ID%22%3A%228362535%22%2C%22NAME%22%3A%22services%22%2C%22GROUP%22%3A%22B%22%7D%2C%22rr_basket%22%3A%7B%22ID%22%3A%228362213%22%2C%22NAME%22%3A%22rr_basket%22%2C%22GROUP%22%3A%22%22%7D%2C%22kt_left%22%3A%7B%22ID%22%3A%228107330%22%2C%22NAME%22%3A%22kt_left%22%2C%22GROUP%22%3A%22B%22%7D%2C%22credit_new_widget%22%3A%7B%22ID%22%3A%228324699%22%2C%22NAME%22%3A%22credit_new_widget%22%2C%22GROUP%22%3A%22A%22%7D%2C%22cartpopap%22%3A%7B%22ID%22%3A%228270671%22%2C%22NAME%22%3A%22cartpopap%22%2C%22GROUP%22%3A%22B%22%7D%2C%22alter_sort%22%3A%7B%22ID%22%3A%228377079%22%2C%22NAME%22%3A%22alter_sort%22%2C%22GROUP%22%3A%22A%22%7D%2C%22dates%22%3A%7B%22ID%22%3A%227895663%22%2C%22NAME%22%3A%22dates%22%2C%22GROUP%22%3A%22%22%7D%2C%22anyquery%22%3A%7B%22ID%22%3A%228402103%22%2C%22NAME%22%3A%22anyquery%22%2C%22GROUP%22%3A%22A%22%7D%7D; dSesn=8ff576f8-67a9-186a-df1d-9e4efac4a33c; iwaf_fingerprint=55c49ce9f14c551ca02458bb6e3ca0b9',
        'origin': 'https://hoff.ru',
        'referer': 'https://hoff.ru/',
        'sec-ch-ua': '"Chromium";v="102", "Opera";v="88", ";Not A Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.115 Safari/537.36 OPR/88.0.4412.53',
        'x-requested-with': 'XMLHttpRequest',
    }

    json_data = {
        'articul': art
    }

    try:
        response = requests.post('https://hoff.ru/vue/search_titles_articul/', cookies=cookies, headers=headers,
                                 json=json_data).json()
        id_product = response.get('data').get('items')[0].get('id')
        articul = response.get('data').get('items')[0].get('articul')

        params = {
            'id': id_product,
            'articul': articul,
        }
        response = requests.get('https://hoff.ru/vue/catalog/product/',
                                params=params, cookies=cookies, headers=headers).json()
        name = response.get('data').get('name')

        box = response.get('data').get('characteristic_tab').get('packages')
        if not box:
            box = ['1']
        data = {
            'name': name,
            'box': len(box)
        }
        return data
    except Exception as ex:
        logger.debug('Товара нет в 6.1 и на сайте =( {}'.format(ex))


if __name__ == '__main__':
    parse('80451234')

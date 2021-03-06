import urllib3
import random

url_path = 'http://exam.sac.net.cn/pages/registration/train-line-register!gsUDDIsearch.action'
image_head = 'http://exam.sac.net.cn/photo/images/'
header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.108 Safari/537.36',
    'X-Forwarded-For': '114.248.238.236'
}

http = urllib3.PoolManager(timeout=10.0)

global null
null = ''


# error = '{'result': 'success', 'message': '访问过于频繁，请稍候访问！'}'

def generateIP():
    ip = ''
    ip += str(random.randint(0, 250)) + '.'
    ip += str(random.randint(0, 250)) + '.'
    ip += str(random.randint(0, 250)) + '.'
    ip += str(random.randint(0, 250))
    return ip


def get_realname(name):
    posts = ['证券经纪有限责任公司', '投资证券有限责任公司', '证券投资咨询有限公司', '证券集团',
             '证券研究所有限责任公司', '投资顾问有限公司',
             '证券有限责任公司', '证券股份有限公司', '证券有限公司',
             '集团股份有限公司', '投资咨询有限公司', '证券(香港)有限公司',
             '有限责任公司', '股份有限公司', '有限公司', '证券公司', '证券']
    for post in posts:
        if name.endswith(post):
            return name[:-len(post)]
    return name


def equal_company(a, b):
    real_a = get_realname(a)
    real_b = get_realname(b)

    return real_a == real_b


def is_person(personID, company):
    param = {
        'filter_EQS_RH#RPI_ID': personID,
        'sqlkey': 'registration',
        'sqlval': 'SEARCH_LIST_BY_PERSON'
    }
    header['X-Forwarded-For'] = generateIP()

    res = http.request('POST', url_path, fields=param, headers=header)
    res_data = eval(res.data.decode())

    if res_data:
        for one in res_data:
            if equal_company(one['AOI_NAME'], company):
                return True
    return False


def search_way(company, name, way):
    param = {
        'filter_EQS_PPP_NAME': name,
        'sqlkey': 'registration',
        'sqlval': way
    }

    header['X-Forwarded-For'] = generateIP()

    res = http.request('POST', url_path, fields=param, headers=header)
    res_data = eval(res.data.decode())

    if res_data:
        for one in res_data:
            if equal_company(one['AOI_NAME'], company):
                return res_data, one
    return res_data, None


def search(company, name):
    total = []

    one, result = search_way(company, name, 'SEARCH_FINISH_NAME')
    if one:
        total += one

    if not result:
        two, result = search_way(company, name, 'SEARCH_FINISH_OTHER_NAME')
        if two:
            total += two

    if not result and len(total) > 0:
        if len(total) == 1:
            result = total[0]
        else:
            for person in total:
                person_id = getPersonID(person['PPP_ID'])
                if is_person(person_id, company):
                    result = person
                    result['personID'] = person_id
                    break

    return result


def getPersonID(PPP_ID):
    param = {
        'filter_EQS_PPP_ID': PPP_ID,
        'sqlkey': 'registration',
        'sqlval': 'SD_A02Leiirkmuexe_b9ID'
    }
    header['X-Forwarded-For'] = generateIP()

    res = http.request('POST', url_path, fields=param, headers=header)
    res_data = eval(res.data.decode())

    if res_data:
        return res_data[0]['RPI_ID']
    else:
        return None


def getImagePath(personID):
    param = {
        'filter_EQS_RPI_ID': personID,
        'sqlkey': 'registration',
        'sqlval': 'SELECT_PERSON_INFO'
    }
    header['X-Forwarded-For'] = generateIP()

    res = http.request('POST', url_path, fields=param, headers=header)
    res_data = eval(res.data.decode())

    if res_data:
        return image_head + res_data[0]['RPI_PHOTO_PATH']
    else:
        return None


if __name__ == '__main__':
    print(equal_company('我你他证券股份有限公司', '我你他有限公司'))

import random
import requests
import json
import re
import time
import xlwt
import jieba

def remove(words):
    #停用词
    stopword = [',', '.', '!', '*', '~', '(', ')', '。', '，', '！', ':', '：', "'", ' ', '`', '?', '@']
    final = ''
    for w in words:
        if w not in stopword:
            final += w
    return final

def writeExcel(workbook, worksheet, x, y, data):
    # 往表格写入内容
    worksheet.write(x, y, data)
    # 保存
    workbook.save("jd.xls")



def main():
    # 创建新的workbook（其实就是创建新的excel）
    workbook = xlwt.Workbook(encoding='ascii')
    # 创建新的sheet表
    worksheet = workbook.add_sheet("data")

    for page in range(0, 10):
        header = {
            'refer': 'https://item.jd.com/',
            'cookie': '',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36 Edg/110.0.1587.50'
        }

        productId = '4044691'

        parm = {
            'callback': 'fetchJSON_comment98',
            'productId': productId,
            'score': '0',
            'sortType': '5',
            'page': page,
            'pageSize': '10',
            'isShadowSku': '0',
            'fold': '1'
        }
        url = 'https://club.jd.com/comment/productPageComments.action'
        res = requests.get(url, params=parm, headers=header)

        print('第%d页正在爬取' % (page + 1))

        # 爬取完成后，需要对页面进行编码，不影响后期的数据提取和数据清洗工作。
        # 使用正则对数据进行提取，返回字符串。
        # 字符串转换为json格式数据。
        res.encoding = 'gb18030'
        html = res.text
        data = re.findall('fetchJSON_comment98\((.*?)\);', html)
        data = json.loads(data[0])  # 将处理的数据进行解析
        comments = data['comments']

        for index, comment in enumerate(comments):
            score = comment['score']
            creationTime = comment['creationTime']
            content = comment['content']
        
        #将商品评分、评价时间、评价内容写入 excel 中
            writeExcel(workbook, worksheet, page * 10 + index, 0, score)
            writeExcel(workbook, worksheet, page * 10 + index, 1, creationTime)
            writeExcel(workbook, worksheet, page * 10 + index, 2, content)
            
            print(content)

            content2 = remove(content)

            #利用 jieba 进行分词
            words = jieba.cut(content2, cut_all=False)

            #将分词内容写入 excel 中
            for index2, word in enumerate(words):
                writeExcel(workbook, worksheet, page * 10 + index, index2 + 3, str(word))

        # 程序休眠
        time.sleep(random.randint(10, 20) * 0.1)

    print("爬取完毕")


if __name__ == '__main__':
    main()



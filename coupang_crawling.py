from bs4 import BeautifulSoup as bs
from typing import Optional, Union, Dict, List
from openpyxl import Workbook
import os
import re
import sys
import requests as rq
import json
from bs4 import BeautifulSoup
import requests

sys.stdout.reconfigure(encoding = 'utf-8')


prod_real_name = ''

def fetch_product_title(url):
    # 요청을 보내고 응답을 받음
    response = requests.get(url)
    if response.status_code == 200:

        soup = BeautifulSoup(response.text, 'html.parser')
        # 특정 클래스를 가진 h2 태그 찾기
        title_tag = soup.find('h2', class_='prod-buy-header__title')
        if title_tag:
            # 텍스트 추출 및 반환
            return title_tag.text.strip()
    return None

def get_headers(key: str, default_value: Optional[str] = None) -> Dict[str, Dict[str, str]]:
    """Get Headers"""
    JSON_FILE = 'json/headers.json'
    json_file_path = os.path.join(os.path.dirname(__file__), JSON_FILE)  # 스크립트의 위치에 따른 절대 경로를 구성합니다.
    try:
        with open(json_file_path, 'r', encoding='UTF-8') as file:
            headers = json.loads(file.read())
        return headers[key]
    except FileNotFoundError:
        if default_value:
            return default_value
        raise EnvironmentError(f"Set the {key} in {json_file_path}")


class Coupang:
    @staticmethod
    def get_product_code(url: str) -> str:
        """Extract PRODUCT CODE from given URL"""
        return url.split('products/')[-1].split('?')[0]

    def __init__(self) -> None:
        self.__headers = get_headers(key='headers')
        

    def main(self, review_url: str, total_reviews_count: int) -> None:
        prod_code = self.get_product_code(review_url)
        self.__headers['referer'] = review_url

        all_reviews = []
        page = 1

        while len(all_reviews) < total_reviews_count:
            url = f'https://www.coupang.com/vp/product/reviews?productId={prod_code}&page={page}&size=10&sortBy=ORDER_SCORE_ASC&ratings=&q=&viRoleCode=3&ratingSummary=true'
            #print(url)
            reviews = self.fetch(url)
            
            if not reviews:
                break  # No more reviews
            all_reviews.extend(reviews)
            if len(reviews) < 10:
                break  # Last page
            page += 1

        self.save_file(all_reviews[:total_reviews_count])



    def fetch(self, url: str) -> List[Dict[str, Union[str, int]]]:
        save_data = []
        with rq.Session() as session:
            response = session.get(url, headers=self.__headers)

            soup = bs(response.text, 'html.parser')
            articles = soup.select('article.sdp-review__article__list')
          
            for article in articles:
                dict_data : Dict[str,Union[str,int]] = dict()
                user_name = article.select_one('span.sdp-review__article__list__info__user__name')
                user_name = '-' if not user_name else user_name.text.strip()
                rating = article.select_one('div.sdp-review__article__list__info__product-info__star-orange')
                rating = 0 if not rating else int(rating.attrs['data-rating'])
                prod_name = article.select_one('div.sdp-review__article__list__info__product-info__name')
                prod_name = '-' if not prod_name else prod_name.text.strip()
                headline = article.select_one('div.sdp-review__article__list__headline')
                headline = 'No headline' if not headline else headline.text.strip()
                review_content = article.select_one('div.sdp-review__article__list__review > div')
                review_content = 'No content' if not review_content else re.sub('[\n\t]', '', review_content.text.strip())
                answer = article.select_one('span.sdp-review__article__list__survey__row__answer')
                answer = 'No answer' if not answer else answer.text.strip()


                dict_data['prod_name'] = prod_name
                dict_data['user_name'] = user_name
                dict_data['rating'] = rating
                dict_data['headline'] = headline
                dict_data['review_content'] = review_content
                dict_data['answer'] = answer

                save_data.append({
                    'prod_name': prod_name,
                    'user_name': user_name,
                    'rating': rating,
                    'headline': headline,
                    'review_content': review_content,
                    'answer': answer
                })

               # print(save_data , '\n')

        return save_data

    @staticmethod
    def save_file(reviews: List[Dict[str, Union[str, int]]]) -> None:
        reviews_str = reviews
        print(reviews_str)
        if not reviews:
            print("No data to save.")
            return

        
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(['상품상세명','구매자 이름','구매자 평점','리뷰 제목','리뷰 내용','만족도'])

        for review in reviews:
            worksheet.append([
                review['prod_name'], review['user_name'], review['rating'],
                review['headline'], review['review_content'], review['answer']
            ])

        savePath : str = os.path.abspath('쿠팡-상품리뷰-크롤링')
        fileName = reviews[0]['prod_name'] + '.xlsx'
        #fileName = prod_real_name + '.xlsx'
        fullPath = os.path.join(savePath, fileName)  # Construct the full file path
       # print(f"File save_path: {fullPath}")
        if not os.path.exists(savePath):
            os.mkdir(savePath)

        workbook.save(fullPath)  # Save the workbook to the constructed path
        workbook.close()
       # print(f"File saved: {fullPath}")
      
if __name__ == '__main__':
    # Get user input for review URL and total reviews count
    review_url = input("리뷰를 가져올 쿠팡의 경로를 입력하세요 :")
    total_reviews_count_input = input("가져오고싶은 리뷰의 갯수를 입력하세요 : ")

    try:
        total_reviews_count = int(total_reviews_count_input)
    except ValueError:
        print("가져오고싶은 리뷰의 갯수를 정수로 입력하세요 ")
        sys.exit(1)

    coupang = Coupang()
    coupang.main(review_url, total_reviews_count)

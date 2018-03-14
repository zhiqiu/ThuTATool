from selenium import webdriver
import time

driver = webdriver.Chrome()
search_url = 'https://guorn.com/stock/strategy?sid=13397.R.82392652883003'
driver.get(search_url)
time.sleep(0.5)

# 元素不可点击
# Element is not clickable at point
# button = driver.find_element_by_css_selector("#holdings-cnt > ul > li.trade-history ")
# button.click()

# 用js实现点击
driver.execute_script('var q=document.getElementsByClassName("trade-history");q[0].click();')

print(driver.page_source)

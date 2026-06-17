from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.window import WindowTypes
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client
import pyautogui
import time 

def load(driver):
    driver.get("https://omni.top-academy.ru")

    search_input_login = driver.find_element(By.XPATH, "//input[@ng-model='form.username']")
    search_input_login.send_keys("")
    search_input_pass = driver.find_element(By.XPATH, "//input[@ng-model='form.password']")
    search_input_pass.send_keys("")
    search_input_login.submit()

def load_page(driver):
    time.sleep(3)
    try:
        load(driver)
    except :
        while True:
            print("Await")
            driver.quit()
            load(driver)
        
    time.sleep(5)
    
    driver.switch_to.new_window(WindowTypes.TAB)
    driver.get("https://omni.top-academy.ru/#/bind/materials")

def click_materials():
    time.sleep(10)
    x = 75 #540
    y = 615 #345
    pyautogui.moveTo(x, y) 
    pyautogui.click() 
    time.sleep(5)

def find_id_material(driver, i):
    
    try:
        elements = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//md-option[@id='select_option_{i}']"))
        )           
        # find_id = driver.find_element(By.XPATH, )
        elements.click()    
        time.sleep(5)
        return True
    
    except Exception as ex:
        print(ex)
        try:
            new_method = driver.find_element(By.XPATH, "//div[@class='md-text ng-binding']")
            

            print("Elements finded")
            new_method.click()
        except:
            for i in range(3):
                try:
                    elements = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//md-option[@id='select_option_323']"))
                    ) 
                    elements.click()
                    print(f"Элемент найден с {i + 1} попытки")
                    time.sleep(5)
                    return True
                except:
                    print(f"Попытка {i + 1} не удалась")
                    time.sleep(2)
    
def copy_lesson(driver,i):
    arr = []
    find_name_material = driver.find_element(By.XPATH, "//h2[@class='has-link-color zag wp-elements-caae2e1386b1d9e9dfd57ba858ee841f wp-block-post-title has-text-color has-background-color has-background has-transparent-background-color has-large-font-size']").text
    arr.append(find_name_material) 
    for j in arr:
        print(f"Название урока: {j}")
    
    time.sleep(4)
    x = 338 #338
    y = 1210 #790
    pyautogui.moveTo(x,y)
    pyautogui.mouseDown(button="left")
    pyautogui.moveRel(0,1000, duration=10)
    pyautogui.mouseUp(button="left")
    pyautogui.hotkey("ctrl", "c")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    word.Selection.Paste()
    doc.SaveAs(rf"C:\Users\perth\OneDrive\Desktop\materials\{j}{i}.docx")
    
    doc.Close()
    word.Quit()

def new_window_for_lesson(driver):
    
        input_element = driver.find_element(By.ID, "input_266")
        text_from_input = input_element.get_attribute("value")
        driver.switch_to.new_window(WindowTypes.TAB)
        try:
            driver.get(f"{text_from_input}")
        except Exception as ex:
            print("Error get", ex)
        time.sleep(6)
    
def save_lesson_word(driver):
        try:
            div_element = driver.find_element(By.XPATH, "//div[@data-elementor-type='wp-page']")
            text_scrape = driver.execute_script("return arguments[0].innerHTML;", div_element)
            print(f"len {len(text_scrape)}")


        except Exception as ex:
            print("Error", ex)

def find_all_element_button_lesson(driver):
    time.sleep(5)

    
    for i in range(1,19):
        try:
            time.sleep(5)
            find_lesson = driver.find_element(By.XPATH, f"//span[contains(text(), 'Урок {i}')]")
            # print(f"найден {i}")
            
            if find_lesson:
                find_lesson.click()
                time.sleep(3)
                try:
                    find_link_lesson = driver.find_element(By.XPATH, "//span[@class='posts-link ng-scope']")

                    print("find element")
                    find_link_lesson.click()

                    time.sleep(3)
                    print("tested1")
                    
                    x = 1106
                    y = 274
                    pyautogui.moveTo(x,y)
                    pyautogui.click()

                    x = 1050
                    y = 680
                    pyautogui.moveTo(x,y)
                    time.sleep(2)
                    pyautogui.click()
                    # driver.close()
                    time.sleep(10)
                    driver.switch_to.new_window(WindowTypes.TAB)
                    driver.get("https://omni.top-academy.ru/#/bind/materials")
                    
                    time.sleep(10)

                except Exception as ex:
                        # print(ex)
                        print("no find button link")
                        new_window_for_lesson(driver)
                        save_lesson_word(driver)
                        copy_lesson(driver,i)

                        time.sleep(10)
                        driver.switch_to.new_window(WindowTypes.TAB)
                        driver.get("https://omni.top-academy.ru/#/bind/materials")
                        
                        time.sleep(10)
            else:
                print("Error")
        except Exception as ex:
            print(f"Не найден элемент {i} {ex}")
            break

def main():
    driver = webdriver.Chrome()
    load_page(driver)
    click_materials()
    for i in range(10):
        # 300 - номер урока по html
        number_lesson = 300 + i

        find_id_material(driver, number_lesson)
        print("working")
        find_all_element_button_lesson(driver)


if __name__ == "__main__":
    main()

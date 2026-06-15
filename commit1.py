from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.window import WindowTypes
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client
import pyautogui
import time 

def load(driver):
    driver.get("")

    search_input_login = driver.find_element(By.XPATH, "//input[@ng-model='form.username']")
    search_input_login.send_keys("katyrlo_g@t.top-academy.ru")
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
    x = 720 #540
    y = 515 #345
    pyautogui.moveTo(x, y) 
    pyautogui.click() 
    time.sleep(5)

def find_id_material(driver):
        
    try:
        elements = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//md-option[@id='select_option_323']"))
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

def name(driver):
    find_name_material = driver.find_element(By.XPATH, "//h2[@class='has-link-color zag wp-elements-caae2e1386b1d9e9dfd57ba858ee841f wp-block-post-title has-text-color has-background-color has-background has-transparent-background-color has-large-font-size']")
    time.sleep(3)
    # print(find_name_material.text)
    return find_name_material.text
    
def test(driver,i):

    file_name = name(driver)
    time.sleep(4)
    x = 463 #338
    y = 1095 #790
    pyautogui.moveTo(x,y)
    pyautogui.mouseDown(button="left")
    pyautogui.moveRel(0,1000, duration=10)
    pyautogui.mouseUp(button="left")
    pyautogui.hotkey("ctrl", "c")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    word.Selection.Paste()
    doc.SaveAs(rf"C:\Users\perth\OneDrive\Desktop\materials{file_name}{i}.docx")
    doc.Close()

    word.Quit()

    # add = pyperclip.paste()
    # doc = Document()
    # doc.add_paragraph(add)
    # doc.save("lesson_test_v1_0.docx")

def find_all_element_button_lesson(driver):
    time.sleep(5)

    
    for i in range(1,19):
        try:
            find_lesson = driver.find_element(By.XPATH, f"//span[contains(text(), 'Урок {i}')]")
            # print(f"найден {i}")
            if find_lesson:
                try:
                    find_lesson.click()

                    input_element = driver.find_element(By.ID, "input_270")

                    try:
                        
                        text_from_input = input_element.get_attribute("value")
                        driver.switch_to.new_window(WindowTypes.TAB)
                        try:
                            driver.get(f"{text_from_input}")
                        except Exception as ex:
                            print("Error get", ex)
                        time.sleep(6)

                        try:
                            div_element = driver.find_element(By.XPATH, "//div[@data-elementor-type='wp-page']")
                            text_scrape = driver.execute_script("return arguments[0].innerHTML;", div_element)
                            print(f"len {len(text_scrape)}")
                            test(driver,i)
                            # driver.close()
                            time.sleep(10)
                            driver.switch_to.new_window(WindowTypes.TAB)
                            driver.get("https://omni.top-academy.ru/#/bind/materials")
                            
                            time.sleep(10)

                            # doc = Document()
                            # doc.add_paragraph(text_scrape)
                            
                            # file_path = os.path.abspath('урок 1425.docx')
                            # file_size = os.path.getsize('урок 1425.docx')
                            # doc.save('урок 1425.docx')
                            # print(f"peth: {file_path}, size: {file_size}")
                        except Exception as ex:
                            print("Error", ex)
                    
                    except:
                        for i in range(3):
                            input_element = driver.find_element(By.ID, "input_281")
                            text_from_input = input_element.get_attribute("value")
                            print(f"attempt {i}")


                except Exception as ex:
                    print("Error", ex)
            else:
                print("Error")
        except Exception as ex:
            print(f"Не найден элемент {i} {ex}")
            break

def find_all_elements(driver):
    time.sleep(10)
    try:
        i = 0
        while True:
            test = driver.find_elements(By.XPATH, f"//span[@text()='Текст {i}']")
            if test: 
                for element in test:
                    print(element.text)
            else:
                print("Ended")
                break
            i += 1
    except Exception as ex:
        print(ex)

    time.sleep(3)
    driver.quit()

def main():
    driver = webdriver.Chrome()
    load_page(driver)
    click_materials()
    find_id_material(driver)
    find_all_element_button_lesson(driver)

    # test()
    # find_all_elements(driver)

if __name__ == "__main__":
    main()

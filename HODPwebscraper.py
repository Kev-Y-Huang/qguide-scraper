from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv


# CONSTANTS
FIRST_DEPARTMENT = 0
FIRST_CLASS = 0

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(
    "C:\\Users\\kevyh\\Downloads\\chromedriver_win32\\chromedriver.exe", options=options)
driver.maximize_window()
driver.get("https://qreports.fas.harvard.edu/browse/index")

# Wait until we are on the first page of the Q guide.
element = WebDriverWait(driver, 40).until(
    EC.presence_of_element_located((By.ID, "bluecourses")))
time.sleep(3)

# Get a list of all the deparmtent buttons.
departments = driver.find_elements(
    By.XPATH, "//div[starts-with(@id, 'dept_card')]")

for i in range(FIRST_DEPARTMENT, len(departments)):  # Iterate over the department buttons.
    department = departments[i]

    # Gets the title element of the current department button.
    department_title = department.find_element(
        By.CSS_SELECTOR, "div[id*='_head']").find_element(By.TAG_NAME, "b").text

    click_failed = True
    clicks = 0
    driver.execute_script("return arguments[0].scrollIntoView();", department)
    department.click()

    courses = department.find_elements(By.CSS_SELECTOR, "a[name^='FAS']")
    j = -1
    if i == FIRST_DEPARTMENT:
        j = FIRST_CLASS - 1

    while j < len(courses) - 1:
        j += 1
        print(str(i) + " " + str(j))
        course = courses[j]
        title = course.text

        click_failed = True
        clicks = 0
        driver.execute_script("return arguments[0].scrollIntoView();", course)

        # As many clicks as possible are contianed in a try-except statenment contained in a while loop with a counter.
        # If a click fails (e.g. b/c of connectivity issues), it is attempted 99 more times every 5 seconds. If it
        # still fails, then the program ends, provides the index of the class where it failed, and a manual
        # restart by the programmer is necessary. This has not happened so far.
        while(click_failed):
            try:
                course.click()
                driver.switch_to.window(driver.window_handles[1])
                element = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "TOC")))
                click_failed = False
            except:
                driver.switch_to.window(driver.window_handles[0])
                clicks += 1

            if clicks == 100:
                print("(i, j): " + str(i) + ", " + str(j))
                print("Website is taking too long to respond.")
                exit()

        try:
            responses = driver.find_element(By.ID, "RespCount").text
            enrollment = driver.find_element(By.ID, "InvitedCount").text
        except:
            responses = "None"
            enrollment = "None"
        try:
            rating = driver.find_element(
                By.XPATH, "/html/body/article/div[3]/div[1]/table/tbody/tr[1]/td[7]").text
        except:
            rating = "None"
        try:
            instructor_rating = driver.find_element(
                By.XPATH, "/html/body/article/div[5]/div[1]/table/tbody/tr[1]/td[7]").text
        except:
            instructor_rating = "None"
        try:
            workload_mean = driver.find_element(
                By.XPATH, "/html/body/article/div[6]/div[2]/div/div[3]/table/tbody/tr[3]/td").text
            workload_std = driver.find_element(
                By.XPATH, "/html/body/article/div[6]/div[2]/div/div[3]/table/tbody/tr[6]/td").text
        except:
            workload_mean = "None"
            workload_std = "None"
        try:
            recommendation = driver.find_element(
                By.XPATH, "/html/body/article/div[7]/div[2]/div/div[3]/table[2]/tbody/tr[2]/td[1]").text
        except:
            recommendation = "None"
        try:
            enrollment_tbody = driver.find_element(By.XPATH, "/html/body/article/div[8]/div[2]/div/div/table[2]/tbody")
            enrollment_rows = enrollment_tbody.find_elements(By.TAG_NAME, "tr")

            enrollment_reasons = {}
            for row in enrollment_rows:
                enrollment_reason = row.find_element(By.TAG_NAME, "th").text
                enrollment_count = row.find_element(By.TAG_NAME, "td").text
                enrollment_reasons[enrollment_reason] = enrollment_count
        except:
            enrollment_reasons = "None"
        try:
            comments_tbody = driver.find_element(
                By.XPATH, "//table[contains(@role, 'presentation')]")
            comments = comments_tbody.find_elements(By.TAG_NAME, "td")
            comments_list = []

            for comment in comments:
                comments_list.append(comment.text)
        except:
            comments_list = ["None"]

        # SEQUEL INSERT
        # The first chunk deals with classes.
        # md = """INSERT INTO classes (name, enrollment, overall, workload, department)
        # VALUES (?, ?, ?, ?, ?)"""
        #crsr.execute(cmd, (str(title), str(enrollment), str(rating), str(workload), str(department_title)))
        # connection.commit()
        #crsr.execute("SELECT * FROM classes WHERE name = ?", (str(title),))
        # ans=crsr.fetchall()
        # print(ans)
        with open('q-scores.csv', 'a', newline='', encoding="utf-8") as f:
            thewriter = csv.writer(f)
            thewriter.writerow(["Spring 2022", department_title, str(title), str(workload_mean), str(
                workload_std), str(rating), str(responses), str(enrollment), str(recommendation), str(enrollment_reasons), str(comments_list)])
            f.close()

        driver.close()
        driver.switch_to.window(driver.window_handles[0])

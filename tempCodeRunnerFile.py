    driver.execute_script("arguments[0].click();", bold_button)
    WebDriverWait(driver, 10).until(
        lambda d: not bold_button.is_displayed() or bold_button.get_attribute("aria-pressed") == "true"
    )
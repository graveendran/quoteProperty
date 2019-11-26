from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from datetime import date
from datetime import datetime

from shutil import copyfile

import selenium.webdriver.support.ui as ui
import selenium.webdriver.support.expected_conditions as EC
import os
import time
import csv
import config
import logging
import logging.handlers

# https://www.python.org/downloads/
# pip install selenium
# https://chromedriver.storage.googleapis.com/index.html?path=76.0.3809.126/
# Store the chromedriver in the same directory as the program file

class harmony():
	
	def __init__(self):
		options = webdriver.ChromeOptions()
		options.add_argument('--ignore-certificate-errors')
		options.add_argument('--ignore-ssl-errors')
		dir_path = os.path.dirname(os.path.realpath(__file__))
		chromedriver = dir_path + "/chromedriver"
		print(dir_path + "/chromedriver")
		os.environ["webdriver.chrome.driver"] = chromedriver
		self.driver = webdriver.Chrome(chrome_options=options, executable_path=chromedriver)
	
	def waitForLoading(self):
		time.sleep(5)
		
	def login(self, url, username, password):
		self.driver.get(url)
		logging.info('################Login into System -- %s #########################' %(username))
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.NAME, "username")))
		element.send_keys(username)
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.NAME, "password")))
		element.send_keys(password)
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.CLASS_NAME, "auth0-label-submit")))
		element.click()
		
	def getElementById(self, id):
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "policyHolders[0].firstName")))
		return element
	
	def gotoHarmony(self, url, addressToQuote, firstName, lastName, agencyEmail, phoneNumber, effectiveDate, shareName, shareEmail):
		logging.info('################Search Address -- %s #########################' %(addressToQuote))
		self.driver.get(url + "search/address")
		product = Select(ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "product"))))
		product.select_by_index(1)
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "address")))
		element.send_keys(addressToQuote)
		element.send_keys(Keys.RETURN)
		time.sleep(2)
		
		logging.info('################Click First Address#########################')
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".survey-wrapper ul li")))
		element.click()
		
		logging.info('################Fill Contact Info#########################')
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "policyHolders[0].firstName")))
		element.send_keys(firstName)
		
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "policyHolders[0].lastName")))
		element.send_keys(lastName)
		
		# element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "policyHolders[0].emailAddress")))
		# element.send_keys(agencyEmail)
		
		# element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "policyHolders[0].primaryPhoneNumber")))
		# element.send_keys(phoneNumber)
		
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "effectiveDate")))
		element.send_keys(effectiveDate)
		try:
			agentCode = Select(ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "agentCode"))))
			# mySelect = Select(self.driver.find_element_by_id("agentCode"))
			agentCode.select_by_index(1)
		except NoSuchElementException:
			try:
				agentCode.select_by_visible_text("WALLY WAGONER")
			except NoSuchElementException:
				element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "agentCode")))
				element.send_keys(Keys.DOWN)

		logging.info('################Submit Contact Info#########################')
		self.driver.find_element_by_css_selector('.btn.btn-primary').click()
		time.sleep(5)

		logging.info('################Fill Underwriting Questions#########################')
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "underwritingAnswers.rented.answer")))
		
		element = self.driver.find_element_by_xpath('//span[@data-test="underwritingAnswers.rented.answer_Never"]')
		self.driver.execute_script("arguments[0].click();", element)
		
		element = self.driver.find_element_by_xpath('//span[@data-test="underwritingAnswers.previousClaims.answer_No claims ever filed"]')
		self.driver.execute_script("arguments[0].click();", element)
		
		element = self.driver.find_element_by_xpath('//span[@data-test="underwritingAnswers.monthsOccupied.answer_10+"]')
		self.driver.execute_script("arguments[0].click();", element)
		
		try:
			element = self.driver.find_element_by_xpath('//span[@data-test="underwritingAnswers.fourPointUpdates.answer_Yes"]')
			self.driver.execute_script("arguments[0].click();", element)
		except NoSuchElementException:
			logging.error('################ Not Found: UW Four Point Updates #########################')

		try:
			element = self.driver.find_element_by_xpath('//span[@data-test="underwritingAnswers.business.answer_No"]')
			self.driver.execute_script("arguments[0].click();", element)
			logging.info('################ UW No Business #########################')
		except NoSuchElementException:
			logging.error('Not Found: UW No Business')

		logging.info('################Submit Underwriting Answers#########################')
		self.driver.find_element_by_css_selector('.btn.btn-primary').click()
		time.sleep(5)

		logging.info('################Submit Coverage Details #########################')
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="QuoteWorkflow"]/div[13]/div')))
		element.click()
		time.sleep(5)
		self.driver.find_element_by_css_selector('.btn.btn-primary').click()
		time.sleep(5)
		logging.info('################ Click Submit Button Coverage #########################')
		element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="QuoteWorkflow"]/div[28]/div/button')))
		element.click()
		# time.sleep(5)
		# self.driver.find_element_by_css_selector('.btn.btn-primary').click()
		try:
			logging.info('################ Submit Share Quote Summary #########################')
			element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="QuoteWorkflow"]/div[1]/button[2]')))			
			element = self.driver.find_element_by_xpath("//form[@id='QuoteWorkflow']/div[1]/button[1]")
			element.click()
			logging.info('################ Click Share Quote Button #########################')
			# time.sleep(5)
			logging.info('################ Fill Share Infor #########################')
			element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "toName")))
			element.send_keys(shareName)
			
			element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.ID, "toEmail")))
			element.send_keys(shareEmail)
			
			logging.info('################ Press Enter to Submit #########################')
			buttons = self.driver.find_element_by_xpath("//*[contains(text(), 'Send Email')]")
			logging.info('################ Click Send Email Button ########################')
			buttons.click()
			logging.info('################ Clicked Send Email Button ########################')
			
			# element = ui.WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//form[@id="sendQuoteSummary"]/div[2]/button[2]')))
			logging.info('################ Click Send Email Button ########################')
			# element.click()
		except TimeoutException:
			logging.error('Not Found: Share Quote Button', element, addressToQuote)
		except NoSuchElementException:
			logging.error('Not Found: Share Quote Button', element, addressToQuote)
		except:
			logging.error('Error: Share Quote Button', element, addressToQuote)			
		
	def shutdown(self):
		self.driver.close()
		
	def readFile(self, filePath):
		with open(filePath, 'rb') as f:
			reader = csv.reader(f)
			addressList = list(reader)
			return addressList
	
	def readAddress(self, filePath):
		with open(filePath) as csvfile:
			reader = csv.DictReader(csvfile)
			return list(reader)
	
	def backupFile(self, filePath):
		# Pick the file name
		parts = os.path.split(filePath)
		fileName = parts[1]

		backupDir = parts[0] + "\\backup"

		if not os.path.exists(backupDir):
			os.makedirs(backupDir)

		copyfile(filePath, backupDir + "\\" + fileName)

	def getEffectiveDate(self, saledate):
		effDate = saleDate.split("/")
		
		if len(effDate[1]) == 1:
			day = "0" + effDate[1]
		else:
			day = effDate[1]

		effDate = effDate[0] + "/" + day + "/" + "2019"

		todayDate = date.today().strftime('%m/%d/%Y')
		compareTodayDate = datetime.strptime(todayDate, '%m/%d/%Y')
		compareEffDate = datetime.strptime(effDate, '%m/%d/%Y')
		# effDate = effDate.strptime('%m/%d/%Y')
		if (compareEffDate >= compareTodayDate):
			return effDate
		else:
			return todayDate
	
	def buildPropertyAddress(self, streetNumber, streetPrefix, streetName, streetUnit, city, zip):
		address = ' '.join([
						row['SiteStreetNumber'], 
						row['SiteStreetPrefix'], 
						row['SiteStreetName'],
						row['SiteStreetUnit'], 
						row['SiteCity'], 
						row['SiteZip']
						]).replace("    ", " ").replace("  ", " ")
		return address
	
	def logConfiguration(self):
		folderPath = os.path.dirname(__file__)
		logFolder = '%s\\log\\' %(folderPath)
		if not os.path.exists(logFolder):
			os.makedirs(logFolder)

		logFileName = 'quoteLogs-%s.log' %(datetime.now().strftime("%Y%m%d-%H%M%S"))
		logging.basicConfig(filename = logFolder + logFileName, 
											level = logging.INFO,
											format = '%(asctime)s:%(levelname)s:%(message)s')

if __name__ == "__main__":
	
	obj = harmony()
	obj.logConfiguration()
	logging.info("****************STARTING QUOTING********************")
	# quit()
	
	url = config.webdetails['url']
	username = config.webdetails['username']
	password = config.webdetails['password']
	filePath = config.filepath
	agencyEmail = config.userdetails['agencyemail']
	phoneNumber = config.userdetails['phonenumber']
	shareName = config.userdetails['sharename']
	shareEmail = config.userdetails['shareemail']

	obj.backupFile(filePath)
	
	listOfAddr = obj.readAddress(filePath)
	obj.login(url, username, password)
	logging.info("****************WAITING TO SIGN-IN********************")
	obj.waitForLoading()
	logging.info("****************SIGNED IN********************")
	logging.info("****************READING THE PROPERTY ADDRESS LIST********************")
	for row in listOfAddr:
		address = obj.buildPropertyAddress(row['SiteStreetNumber'], row['SiteStreetPrefix'], 
								row['SiteStreetName'], row['SiteStreetUnit'], row['SiteCity'], row['SiteZip'])
		firstName = row['First Name']
		lastName = row['Last Name']
		saleDate = row['SaleDate']

		effectiveDate = obj.getEffectiveDate(saleDate)
		
		logging.info("****************QUOTING********************")
		logging.info(address, firstName, lastName, effectiveDate)

		obj.gotoHarmony(url, address, firstName, lastName, agencyEmail, phoneNumber, effectiveDate, shareName, shareEmail)
	 	obj.waitForLoading()
	
	obj.shutdown()
	logging.info('#######################End########################')

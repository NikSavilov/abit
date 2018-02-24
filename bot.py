import requests, time, os, httplib2, random, base64, zipfile, pdfkit
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')
    print(credential_path)
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def mail_parser(mail='elen-savilova@yandex.ru', amount = 1):
	credentials = get_credentials()
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('gmail', 'v1', http=http)
	f_list = []
	mes_list = service.users().messages().list(maxResults = amount, userId = 'me', q = 'from:' + mail).execute()
	for message in mes_list['messages']:
		full_message = service.users().messages().get(id = message['id'], userId = 'me').execute()
		for part in full_message['payload']['parts']:
			if part['mimeType'] == 'application/zip':
				attachment = service.users().messages().attachments().get(
					id = part['body']['attachmentId'], 
					messageId = full_message['id'], 
					userId = 'me').execute()
				data = attachment['data']
				name = part['filename']
				file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
				file = open('loaded/' + name, 'wb')
				f_list.append(name)
				file.write(file_data)
				file.close()
	return f_list

def unzip(files_list):
	unzip_list = []
	for name in files_list:
		zip_pack = zipfile.ZipFile('loaded/' + name)
		for f in zip_pack.namelist():
			print('unzipping:  ' + f)
			_ , ext = os.path.splitext(f)
			if ext == '.xml':
				zip_pack.extract(path = 'loaded/unzip', member = f)
				unzip_list.append('loaded/unzip/' + f)
		zip_pack.close()		
	return unzip_list
		
def load_image(unzip_list):
	url = "https://rosreestr.ru/wps/portal/p/cc_ib_portal_services/cc_vizualisation/"
	chromeOptions = webdriver.ChromeOptions() 
	prefs = { 
	"browser.helperApps.alwaysAsk.force" :False, 
	"browser.helperApps.neverAsk.saveToDisk" : "application/pdf;text/plain;text/csv"} 
	chromeOptions.add_experimental_option("prefs",prefs)
	driver = webdriver.Chrome('chromedriver.exe')
	for path in unzip_list:
		driver.get(url) 
		time.sleep(1)
		upload_box = driver.find_element(By.ID,"xml_file");
		upload_box.send_keys('C:/code/abit/' + path)
		driver.find_element(By.CLASS_NAME,"terminal-button-bright").click()
		link_load = driver.find_element(By.XPATH, '/html/body/div/section/div[1]/table/tbody/tr/td/div[2]/div/div/div/div/table/tbody/tr[3]/td/form/table/tbody/tr[3]/td[2]/table/tbody/tr/td/table/tbody/tr/td[2]/a')
		link_load.click()
		driver.switch_to.window(driver.window_handles[-1]);
		page = driver.page_source
		soup = BeautifulSoup(page,'html5lib')
		

		name , _ = os.path.splitext(path)
		pdfkit.from_string(page,name + '.pdf')


		
files = mail_parser()
unzip_list = unzip(files)
load_image(unzip_list)
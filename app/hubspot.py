import time
import requests 
import xlrd 
import json 
from datetime import datetime
import os


# Hubspot API Key 
hapiKey = "fca86e8c-f52f-46ab-a9e4-c2354e39d2f7"

def zuweisungen(file_path):
   print("script called",file_path)
   state = 2     
   # * Static variables       
   col_num_patient_nr = 1 
   col_num_zuweisungs_art = 0

   # ** Function to get last row in Excel Sheet                
   def get_first_row():
      first_found = False
      first = None
         
      for i in range(0,sheet.nrows-1) : 
            if not first_found:                                    
                  if(sheet.cell_value(i,col_num_zuweisungs_art)=="Bemerkung: Wie haben Sie von uns erfahren") : 
                        first = i+2
                        first_found = True
                        break
                  else: 
                        continue  
                        
            else:
               break                                                          
      return first
   # ** Function to get last row in Excel Sheet
   def get_last_row():
      last = None 
      temp_last = None
      last_found = False
      for i in range(sheet.nrows-1,1,-1) : 
            if not last_found:                                    
               for j in range (sheet.ncols) :                          
                  if(sheet.cell_value(i-1,j)=="") : 
                        temp_last = i
                        continue 
                  else: 
                        last=temp_last  
                        last_found = True
                        break
            else:
               break                                                          
      return last 

   
   def update_Deal(dealid, zuweisungs_art):
      url = "https://api.hubapi.com/crm/v3/objects/deals/"+str(dealid)
      querystring = {"hapikey":hapiKey}
      #payload = "{\"properties\":{\"dealname\":\""+newdealname+"\",\"patient\":\""+str(patientnr)+"\"}}"
      payload = "{\"properties\":{\"zuweisungsart\":\""+zuweisungs_art+"\"}}"
      headers = {
         'accept': "application/json",
         'content-type': "application/json"
         }

      response = requests.request("PATCH", url, data=payload.encode('utf-8'), headers=headers, params=querystring)
      
      if response.status_code == 200:
         data = response.json()
         return True
      else: 
         response = requests.request("PATCH", url, data=payload.encode('utf-8'), headers=headers, params=querystring)
         if response.status_code == 200:
            data = response.json()
            return True
         else:
            return False 

   def find_deal_patientnr(patient_nr):
      deals=[]
      url = "https://api.hubapi.com/crm/v3/objects/deals/search?hapikey="+hapiKey
      headers={}
      headers["Content-Type"]="application/json"
      params = json.dumps({
            "filterGroups":[
            {
               "filters":[
               {
                  "propertyName": "patient",
                  "operator": "EQ",
                  "value": patient_nr
               }
               ]
               }
            ]
            })

      response = requests.post(url=url,headers = headers, data = params)
      
      if response.status_code == 200: 
         data = response.json()
         if len(data["results"]) == 0: 
            return False  
         else: 
            for y in range(0,len(data["results"])):
               deals.append(data["results"][y]["id"])
            return deals 
      else: 
         response = requests.post(url=url,headers = headers, data = params)
         if response.status_code == 200: 
            data = response.json()
            if len(data["results"]) == 0: 
               return False  
            else: 
               for y in range(0,len(data["results"])):
                  deals.append(data["results"][y]["id"])
               return deals
               
         else:
            return False
   # Select Exce File and open it 
   
   workbook = xlrd.open_workbook(file_path)
   sheet = workbook.sheet_by_index(0)
   rows = sheet.nrows
   
   # Find last row with elements in the selected excel file 
   last_row = get_last_row()-1
   first_row = get_first_row()

   # Loop trough each row
   for row in range(first_row,last_row):
      # Map the corresponding properties from actual row to local variables     
      zuweisung_patient_nr = int(sheet.cell(row,col_num_patient_nr).value)
      zuweisung_art = sheet.cell(row,col_num_zuweisungs_art).value
      print("Patient Nr: ", zuweisung_patient_nr)
      # Find the corresponding Hubspot Deals with the patient nr  
      deals = find_deal_patientnr(zuweisung_patient_nr)
      if deals == False: 
         print("No Deal with this Patient Nr. is existing")
         actual = row - (first_row-1)  
         status_last = last_row -(first_row-1)
         bias = (100/status_last)
         status_actual = round(actual * bias,1) 
         yield "data:" + str(status_actual) + "\n\n"        
                        
         
      else:
         # Get amount of deals 
         amount_deals = len(deals)
         # take Zuweisungsart 

         # Update Deal 
         for x in range(0,amount_deals):
               update = update_Deal(dealid= deals[x],zuweisungs_art=zuweisung_art)
               if update == True:
                     print("Deal with ", deals[x], "succesful updatet with Zuweisungsart ", zuweisung_art)
               else:
                     print("Deal can't updatet")                    
         actual = row - (first_row-1) 
         status_last = last_row -(first_row-1)
         bias = (100/status_last)
         status_actual = round(actual * bias,1)  
         yield "data:" + str(status_actual) + "\n\n"     
         
                              
      if actual == status_last-1:
            print("delete file")
            os.remove(os.path.join(file_path))
            actual = row - (first_row-1) 
            status_last = last_row -(first_row-1)
            bias = (100/status_last)
            status_actual = round((actual +1) * bias,1)       
            yield "data:" + str(status_actual) + "\n\n"
                     
            break


def aerzte_import(file_path):
      print("script called",file_path)
      state = 2      
      # Funktion für API Load     
      # * Static variables 
      col_num_mail = 10
      col_num_firstname = 2
      col_num_lastname = 1
      col_num_fachtitel = 3
      col_num_phone = 11
      col_num_street = 4
      col_num_city = 6
      col_num_zip = 5
      col_num_gln = 8
      col_num_fax = 12
      col_num_country = 7

      col_num_patient = 17
      col_num_closedate = 13 
      col_num_gesetz = 15
      col_num_betrag = 16

      error_dict = {"Mail":[],"gln": [], "Error Text":[],"Response Status Code":[]};

      ###### Here are the Functions defined 
      # ** Function to get the first row in Excel Sheet
      def get_first_row():
         first_found = False
         first = None
         temp_first = 0 
         for i in range(1,sheet.nrows-1) : 
               if not first_found:                                    
                     if(sheet.cell_value(i,0)=="Zuweisender Arzt: Firma") : 
                           first = i+1
                           first_found = True
                           break
                     else: 
                           continue  
                           
               else:
                  break                                                          
         return first 

      # ** Function to get last row in Excel Sheet
      def get_last_row():
         last = None 
         last_found = False
         temp_last = 0
         for i in range(sheet.nrows-1,1,-1) : 
               if not last_found:                                    
                  for j in range (sheet.ncols) :                          
                     if(sheet.cell_value(i-1,j)=="") : 
                           temp_last = i
                           continue
                     else: 
                           last=temp_last  
                           last_found = True
                           break
               else:
                  break                                                          
         return last 

                     
      # ** Function to check if a deal still exists in Hubspot 
      def search_Deal_Contact(amount, dealname, closedate):
         url_search_contact = "https://api.hubapi.com/crm/v3/objects/deals/search?hapikey="+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         params = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "amount",
                     "operator": "EQ",
                     "value": amount
                  },
                  {
                     "propertyName": "dealname",
                     "operator": "EQ",
                     "value": dealname
                  }, 
                  {
                     "propertyName": "closedate",
                     "operator": "EQ",
                     "value": closedate
                  },
                  ]
                  }
               ]
               })
         if amount ==  "": 
            params = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "dealname",
                     "operator": "EQ",
                     "value": dealname
                  },
                  {
                     "propertyName": "closedate",
                     "operator": "EQ",
                     "value": closedate
                  }
                  ]
                  }
               ]
               })
         
         response = requests.post(url=url_search_contact,headers = headers, data = params)
         if response.status_code == 200: 
            data = response.json() 
            exists = None
            for i in range(0, len(data["results"])): 
                  if i == len(data["results"]) and exists == False:
                     exists =  False           
                  actual_dealname = data["results"][i]["properties"]["dealname"]
                  if actual_dealname == dealname:
                     exists =  True 
                  else: 
                     return False 
            return exists 
         else: 
            response = requests.post(url=url_search_contact,headers = headers, data = params)
            if response.status_code == 200: 
               data = response.json() 
               exists = None
               for i in range(0, len(data["results"])): 
                     if i == len(data["results"]) and exists == False:
                        exists =  False           
                     actual_dealname = data["results"][i]["properties"]["dealname"]
                     if actual_dealname == dealname:
                        exists =  True 
                     else: 
                        return False 
               return exists
            else: 
               return False 
         
      # ** Function to create a Hubspot contact and get vid
      def create_Contact(lastname, firstname, fachtitel, street, zip, city, country, gln, mail, phone, fax):
         url_create_contact = "https://api.hubapi.com/contacts/v1/contact/?hapikey="+hapiKey
         
         headers = {}
         headers["Content-Type"]="application/json"
         raw_data = {"properties": [
                  {
                  "property": "email",
                  "value": mail
                  },
                  {
                  "property": "firstname",
                  "value": firstname
                  },
                  {
                  "property": "lastname",
                  "value": lastname
                  },
                  {
                  "property": "fachtitel",
                  "value": fachtitel
                  },
                  {
                  "property": "gln",
                  "value": gln
                  },
                  {
                  "property": "country",
                  "value": country
                  },
                  {
                  "property": "address",
                  "value": street
                  },
                  {
                  "property": "city",
                  "value": city
                  },
                  {
                  "property": "zip",
                  "value": zip
                  },
                  {
                  "property": "phone",
                  "value": phone
                  },
                  {
                  "property": "fax",
                  "value": fax
                  }
               ]}
         if mail == "":
               del raw_data["properties"][0]              
         
         data = json.dumps(raw_data)

         response = requests.post( url = url_create_contact, data = data, headers = headers )
         if response.status_code == 200: 
            response_data = response.json()
            hs_vid = response_data["vid"]
            return hs_vid
         else: 
            response = requests.post( url = url_create_contact, data = data, headers = headers )
            if response.status_code == 200: 
               response_data = response.json()
               hs_vid = response_data["vid"]
               return hs_vid

            else:
               return False
         

      # ** Function to get the Hubspot contact vid via the GLN
      def search_Contact(gln):
         url_search_contact = "https://api.hubapi.com/crm/v3/objects/contacts/search?hapikey="+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         data = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "gln",
                     "operator": "EQ",
                     "value": gln
                  }
                  ]
                  }
               ]
               })

         response = requests.post(url=url_search_contact,headers = headers, data = data)
                     
         if response.status_code == 200:
               data = response.json() 
               if data["total"] == 0:
                  return False
               else:
                  hs_vid = data["results"][0]["id"]
               return hs_vid
         elif response.status_code == 404:
               print(data)
               print("No contact with the requested gln exists, contact needs to be created first")
               return False
         else: 
            response = requests.post(url=url_search_contact,headers = headers, data = data)
            if response.status_code == 200:
               data = response.json() 
               if data["total"] == 0:
                  return False
               else:
                  hs_vid = data["results"][0]["id"]
               return hs_vid
               
            else: 
               return False         

      # ** Function to get the Hubspot contact vid
      def get_VID(mail):     
         url_get_vid_by_mail = "https://api.hubapi.com/contacts/v1/contact/email/"+mail+"/profile?hapikey="+hapiKey
         
         response = requests.get(url=url_get_vid_by_mail)
         if response.status_code == 200:
               data = response.json()
               hs_vid = data["vid"]
               return hs_vid
         else:
            response = requests.get(url=url_get_vid_by_mail)  
            if response.status_code == 200:
               data = response.json()
               hs_vid = data["vid"]
               return hs_vid
            elif response.status_code == 404:
               print("No contact with the requested email exists, contact needs to be created first")
               return False
            else: 
               return "Error"
      # ** Function to check if the deal is a Erstzuweisung    
      def check_erstzuweisung(patient_nr, deal_closedate):
            deals=[]
            url = "https://api.hubapi.com/crm/v3/objects/deals/search?hapikey="+hapiKey
            headers={}
            headers["Content-Type"]="application/json"
            params = json.dumps({
                  "filterGroups":[
                  {
                     "filters":[
                     {
                        "propertyName": "patient",
                        "operator": "EQ",
                        "value": patient_nr
                     }
                     ]
                     }
                  ],
                  "sorts": [
                    {
                        "propertyName": "closedate",
                        "direction": "ASCENDING"
                    }
                    ]
                  })

            response = requests.post(url=url,headers = headers, data = params)
            if response.status_code == 200: 
                data = response.json()            
                if len(data["results"]) == 0: 
                    return True  
                else: 
                  Erstzuweisung_closedate = data["results"][0]["properties"]["closedate"]
                  if deal_closedate in Erstzuweisung_closedate:
                     return True 
                  else:
                     return False 
                    
            else: 
                response = requests.post(url=url,headers = headers, data = params) 
                if response.status_code == 200: 
                  data = response.json()            
                  if len(data["results"]) == 0: 
                     return True  
                  else: 
                     Erstzuweisung_closedate = data["results"][0]["properties"]["closedate"]
                     if deal_closedate in Erstzuweisung_closedate:
                        return True 
                     else:
                        return False        

      # ** Function to create deal in Hubspot 
      def create_Deal(hs_vid,zip, gln, patient_nr, unix_closedate, gesetz, betrag, dealname, erstzuweisung):
         url_create_deal= 'https://api.hubapi.com/deals/v1/deal?hapikey='+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         data = json.dumps({
               "associations": {
                  "associatedCompanyIds": [],
                  "associatedVids": [
                  hs_vid
                  ]
                  },
               "properties": [
                  {
                  "value": dealname,
                  "name": "dealname"
                  },
                  {
                  "value": "closedwon",
                  "name": "dealstage"
                  },
                  {
                  "value": "default",
                  "name": "pipeline"
                  },
                  {
                  "value": gln,
                  "name": "gln"
                  },
                  {
                  "value": patient_nr,
                  "name": "patient"
                  },
                  {
                  "value": zip,
                  "name": "plz"
                  },
                  {
                  "value": gesetz,
                  "name": "gesetz"
                  },
                  {
                  "value": unix_closedate,
                  "name": "closedate"
                  },
                  {
                  "value": betrag,
                  "name": "amount"
                  },
                  {
                  "value": erstzuweisung,
                  "name": "erstzuweisung"
                  },
                  {
                  "value": "importcuramed",
                  "name": "dealtype"
                  }
               ]
               })

         response = requests.post(url_create_deal, headers = headers, data = data)
         
         if response.status_code == 200:
               data = response.json()
               deal_id = data["dealId"]
               return deal_id
         else: 
            response = requests.post(url_create_deal, headers = headers, data = data)
            if response.status_code == 200:
               data = response.json()
               deal_id = data["dealId"]
               return deal_id
            else:     
               return False


      # Select Exce File and open it 
      workbook = xlrd.open_workbook(file_path)
      sheet = workbook.sheet_by_index(0)
      rows = sheet.nrows
      
      # Find last row with elements in the selected excel file 
      last_row = get_last_row()-1
      first_row = get_first_row()
      # Loop trough each row
      for row in range(first_row,last_row+1):
         # Map the corresponding properties from actual row to local variables     
         customer_lastname = sheet.cell(row,col_num_lastname).value
         customer_firstname = sheet.cell(row,col_num_firstname).value
         customer_fachtitel = sheet.cell(row, col_num_fachtitel).value
         customer_street = sheet.cell(row,col_num_street).value
         customer_zip = sheet.cell(row,col_num_zip).value
         customer_city = sheet.cell(row,col_num_city).value
         customer_country = sheet.cell(row,col_num_country).value
         customer_gln = sheet.cell(row,col_num_gln).value
         customer_mail = sheet.cell(row,col_num_mail).value
         customer_phone = sheet.cell(row,col_num_phone).value
         customer_fax = sheet.cell(row,col_num_fax).value
         customer_patient = int(sheet.cell(row,col_num_patient).value)
         deal_gesetz = sheet.cell(row,col_num_gesetz).value     
         deal_betrag = sheet.cell(row,col_num_betrag).value    

         unix_val = sheet.cell(row,col_num_closedate).value
         unix_close_date = int(((unix_val - 25569)*86400)*1000)
         excel_deal_closedate = xlrd.xldate.xldate_as_datetime(sheet.cell(row,col_num_closedate).value,0)
         # Convert the closedate to Unix
         str_deal_closedate = excel_deal_closedate.strftime("%Y-%m-%d")      
         
         # Find corresponding Hubspot Contact VID from actual row 
         if customer_mail == "":
               # Search for contact via gln 
               customer_vid = search_Contact(customer_gln)
               # If no contact with the gln exists, Contact needs to be created
               if customer_vid == False:
                  customer_vid=create_Contact(customer_lastname, customer_firstname, customer_fachtitel, customer_street, customer_zip, customer_city, customer_country, customer_gln, customer_mail, customer_phone, customer_fax)
               # else go ahead and create deal with this vid 
               else: 
                  pass 
               
         else:
               customer_vid = get_VID(customer_mail)
               if customer_vid == False:
                  customer_vid=create_Contact(customer_lastname, customer_firstname, customer_fachtitel, customer_street, customer_zip, customer_city, customer_country, customer_gln, customer_mail, customer_phone, customer_fax)
                  print("Customer didnt exist & is created", customer_vid)
               else:
                  customer_vid = get_VID(customer_mail)    

         # Create dealname 
         dealname = customer_lastname +"-" + str(customer_patient)+"-" + str_deal_closedate

         # Check if the deal still exists 
         deal_exist = search_Deal_Contact(amount = deal_betrag, dealname = dealname, closedate = unix_close_date)

         # Check if the deal is a Erestzuweisung 
         erstzuweisung = check_erstzuweisung(patient_nr=customer_patient,deal_closedate=str_deal_closedate)
         if deal_exist == True:
            status_actual = row - (first_row-1)
            status_last = last_row - (first_row-1)
            print("Deal still exists")
            print("Contact", status_actual,"/",status_last, "Done") 
            actual = row - (first_row-1) 
            status_last = last_row -(first_row-1)
            bias = (100/status_last)
            status_actual = round(actual * bias,1)  
            yield "data:" + str(status_actual) + "\n\n" 
            continue
         else: 
            # Create deal in Hubspot via API 
            dealnumber = create_Deal(customer_vid,customer_zip, customer_gln, customer_patient, unix_close_date, deal_gesetz, deal_betrag, dealname, erstzuweisung)
            status_actual = row - (first_row-1)
            status_last = last_row - (first_row-1)
            print("Deal created:", dealnumber)
            print("Contact", status_actual,"/",status_last, "Done")
            actual = row - (first_row-1) 
            status_last = last_row -(first_row-1)
            bias = (100/status_last)
            status_actual = round(actual * bias,1)  
            yield "data:" + str(status_actual) + "\n\n"        
                       
      
         if actual == status_last-1:
               print("delete file")
               os.remove(os.path.join(file_path))
               actual = row - (first_row-1) 
               status_last = last_row -(first_row-1)
               bias = (100/status_last)
               status_actual = round((actual +1) * bias,1)       
               yield "data:" + str(status_actual) + "\n\n"
                           
               break


def gruppenpraxen_import(file_path):
      print("script called",file_path)
      state = 2      
      # Funktion für API Load     
      # * Static variables       
      col_num_company = 0 
      col_num_street = 4
      col_num_zip = 5 
      col_num_city = 6
      col_num_country = 7
      col_num_gln = 8
      col_num_mail = 10
      col_num_phone = 11
      col_num_fax = 12 
      col_num_patient = 17    
      
      col_num_closedate = 13
      col_num_gesetz = 15
      col_num_betrag = 16

      
      ###### Here are the Functions defined 
      # ** Function to get the first row in Excel Sheet
      def get_first_row():
         first_found = False
         first = None
            
         for i in range(1,sheet.nrows-1) : 
               if not first_found:                                    
                     if(sheet.cell_value(i,0)=="Zuweisender Arzt: Firma") : 
                           first = i+2
                           first_found = True
                           break
                     else: 
                           continue  
                           
               else:
                  break                                                          
         return first
      # ** Function to get last row in Excel Sheet
      def get_last_row():
         temp_last = None
         last = None
         last_found = False
         for i in range(sheet.nrows-1,1,-1) : 
               if not last_found:                                    
                  for j in range (sheet.ncols) :                          
                     if(sheet.cell_value(i-1,j)=="") : 
                           temp_last = i
                           continue
                     else: 
                           last=temp_last  
                           last_found = True
                           break
               else:
                  break                                                          
         return last 

      
      # Function to search a Deal and check if it exists already
      def search_Deal_Check(amount, company_name, closedate):
         url_search_contact = "https://api.hubapi.com/crm/v3/objects/deals/search?hapikey="+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         params = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "amount",
                     "operator": "EQ",
                     "value": amount
                  }, 
                  {
                     "propertyName": "closedate",
                     "operator": "EQ",
                     "value": closedate
                  },
                  ]
                  }
               ]
               })
         if amount ==  "": 
            params = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "closedate",
                     "operator": "EQ",
                     "value": closedate
                  }
                  ]
                  }
               ]
               })
         
         response = requests.post(url=url_search_contact,headers = headers, data = params)
         print(response)
         dealname_exists = None

         if response.status_code == 200:
            data = response.json() 
            for i in range(0,len(data["results"])):
                  if i == len(data["results"]) and dealname_exists == False:
                     return False
                  len_company_name = len(company_name)
                  company= data["results"][i]["properties"]["dealname"]
                  for i in range(0, len_company_name):
                     if i == len_company_name-3 and dealname_exists == True:
                        dealname_exists = True                     
                        break
                     if  company[i] == company_name[i]:
                        dealname_exists = True
                     else: 
                        dealname_exists = False                      
                        break 
         else:
            response = requests.post(url=url_search_contact,headers = headers, data = params)  
            print(response)
            data = response.json() 
            dealname_exists = None
            for i in range(0,len(data["results"])):
               if i == len(data["results"]) and dealname_exists == False:
                  return False
               len_company_name = len(company_name)
               company= data["results"][i]["properties"]["dealname"]
               for i in range(0, len_company_name):
                  if i == len_company_name-3 and dealname_exists == True:
                     dealname_exists = True                     
                     break
                  if  company[i] == company_name[i]:
                     dealname_exists = True
                  else: 
                     dealname_exists = False                      
                     break 
         print("Deal exists or not: ", dealname_exists)
         return dealname_exists       
         

      # ** Function to create for the company an unique number per deal 
      def create_DealNumber(hs_vid):
         url = "https://api.hubapi.com/crm/v3/objects/companies/"+hs_vid+"/associations/deals"
         querystring = {"paginateAssociations":"false","limit":"500","hapikey":hapiKey}
         headers = {'accept': 'application/json'}

         response = requests.request("GET", url, headers=headers, params=querystring)
         print(response)
         if response.status_code == 200: 
            data = response.json()
            amount_deals = len(data["results"])
            return amount_deals+1
         else:
            response = requests.request("GET", url, headers=headers, params=querystring)   
            print("API Retry:", response)
            data = response.json()
            amount_deals = len(data["results"])
            return amount_deals+1

      # ** Function to create a Hubspot Company and get vid
      def create_Company(company_name, street, zip, city, country, gln, mail, phone, fax):
         url = "https://api.hubapi.com/crm/v3/objects/companies"
         querystring = {"hapikey":hapiKey}
         headers = {
               'content-type': "application/json"
            }
         payload = "{\"properties\":{\"email\":\""+mail+"\",\"name\":\""+company_name+"\",\"address\":\""+street+"\",\"zip\":\""+zip+"\",\"city\":\""+city+"\",\"gln\":\""+gln+"\",\"phone\":\""+phone+"\", \"fax\":\""+fax+"\"}}"
                  
         response = requests.request("POST", url, data=payload.encode('utf-8'), headers=headers, params=querystring)
         
         if response.status_code == 200 or response.status_code == 201:
            response_data = response.json()
            hs_vid = response_data["id"]
            print("Company created, ID:", hs_vid)
            return hs_vid
         
         elif response.status_code == 409:
            print("Company cant created, Deal still exists")
            return False
         else: 
            response = requests.request("POST", url, data=payload.encode('utf-8'), headers=headers, params=querystring)
            if response.status_code == 200 or response.status_code == 201:
               response_data = response.json()
               hs_vid = response_data["id"]
               print("Company created, ID:", hs_vid)
               return hs_vid
            else:   
               return "Error"

      # ** Function to get the Hubspot Company vid via the GLN
      def search_Company(gln):
         url_search_contact = "https://api.hubapi.com/crm/v3/objects/companies/search?hapikey="+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         data = json.dumps({
               "filterGroups":[
               {
                  "filters":[
                  {
                     "propertyName": "gln",
                     "operator": "EQ",
                     "value": gln
                  }
                  ]
                  }
               ]
               })

         response = requests.post(url=url_search_contact,headers = headers, data = data)
         print(response)
         
         
         if response.status_code == 200:
               data = response.json() 
               if data["total"] == 0:
                  return False
               else:
                  hs_vid = data["results"][0]["id"]
               return hs_vid
         elif response.status_code == 404:
               print(data)
               print("No contact with the requested gln exists, contact needs to be created first")
               return False
         else: 
            response = requests.post(url=url_search_contact,headers = headers, data = data)
            print(response) 
            if response.status_code == 200:
               data = response.json() 
               if data["total"] == 0:
                  return False
               else:
                  hs_vid = data["results"][0]["id"]
               return hs_vid 
                        
            else:
               return False   
      # ** Function to check if the deal is a Erstzuweisung    
      def check_erstzuweisung(patient_nr, deal_closedate):
            deals=[]
            url = "https://api.hubapi.com/crm/v3/objects/deals/search?hapikey="+hapiKey
            headers={}
            headers["Content-Type"]="application/json"
            params = json.dumps({
                  "filterGroups":[
                  {
                     "filters":[
                     {
                        "propertyName": "patient",
                        "operator": "EQ",
                        "value": patient_nr
                     }
                     ]
                     }
                  ],
                  "sorts": [
                    {
                        "propertyName": "closedate",
                        "direction": "ASCENDING"
                    }
                    ]
                  })

            response = requests.post(url=url,headers = headers, data = params)
            if response.status_code == 200: 
                data = response.json()            
                if len(data["results"]) == 0: 
                    return True  
                else: 
                  Erstzuweisung_closedate = data["results"][0]["properties"]["closedate"]
                  if deal_closedate in Erstzuweisung_closedate:
                     return True 
                  else:
                     return False 
                    
            else: 
                response = requests.post(url=url,headers = headers, data = params) 
                if response.status_code == 200: 
                  data = response.json()            
                  if len(data["results"]) == 0: 
                     return True  
                  else: 
                     Erstzuweisung_closedate = data["results"][0]["properties"]["closedate"]
                     if deal_closedate in Erstzuweisung_closedate:
                        return True 
                     else:
                        return False      
               
      # ** Function to create deal in Hubspot 
      def create_Company_Deal(hs_vid,zip, gln, unix_closedate, gesetz, betrag, dealname, patient_nr, erstzuweisung):
         url_create_deal= 'https://api.hubapi.com/deals/v1/deal?hapikey='+hapiKey
         headers={}
         headers["Content-Type"]="application/json"
         data = json.dumps({
               "associations": {
                  "associatedCompanyIds": [hs_vid],
                  "associatedVids": []
                  },
               "properties": [
                  {
                  "value": dealname,
                  "name": "dealname"
                  },
                  {
                  "value": "closedwon",
                  "name": "dealstage"
                  },
                  {
                  "value": "default",
                  "name": "pipeline"
                  },
                  {
                  "value": gln,
                  "name": "gln"
                  },
                  {
                  "value": patient_nr,
                  "name": "patient"
                  },
                  {
                  "value": zip,
                  "name": "plz"
                  },
                  {
                  "value": gesetz,
                  "name": "gesetz"
                  },
                  {
                  "value": unix_closedate,
                  "name": "closedate"
                  },
                  {
                  "value": betrag,
                  "name": "amount"
                  },
                  {
                  "value": erstzuweisung,
                  "name": "erstzuweisung"
                  },
                  {
                  "value": "importcuramed",
                  "name": "dealtype"
                  }
               ]
               })

         response = requests.post(url_create_deal, headers = headers, data = data)
         
         if response.status_code == 200:
               data = response.json()
               deal_id = data["dealId"]
               return deal_id
         else: 
            response = requests.post(url_create_deal, headers = headers, data = data)   
            if response.status_code == 200:
               data = response.json()
               deal_id = data["dealId"]
               return deal_id
            else:   
               return False


      # Select Exce File and open it 
      workbook = xlrd.open_workbook(file_path)
      sheet = workbook.sheet_by_index(0)
      rows = sheet.nrows
      
      # Find last row with elements in the selected excel file 
      last_row = get_last_row()-1
      first_row = get_first_row()
      # Loop trough each row
      #changed to 9
      for row in range(first_row,last_row):
         # Map the corresponding properties from actual row to local variables     
         company_name = sheet.cell(row,col_num_company).value
         company_street = sheet.cell(row,col_num_street).value
         company_zip = sheet.cell(row,col_num_zip).value
         company_city = sheet.cell(row,col_num_city).value
         company_country = sheet.cell(row,col_num_country).value
         company_gln = sheet.cell(row,col_num_gln).value
         company_mail = sheet.cell(row,col_num_mail).value
         company_phone = sheet.cell(row,col_num_phone).value
         company_fax = sheet.cell(row,col_num_fax).value
         deal_gesetz = sheet.cell(row,col_num_gesetz).value     
         deal_betrag = sheet.cell(row,col_num_betrag).value 
         company_patient_nr = int(sheet.cell(row,col_num_patient).value)         
      
         # Convert closedate to Unix for the Creation of the Deal 
         unix_val = sheet.cell(row,col_num_closedate).value
         unix_close_date = int(((unix_val - 25569)*86400)*1000)
         # Convert closedate to String for the Dealname
         excel_deal_closedate = xlrd.xldate.xldate_as_datetime(sheet.cell(row,col_num_closedate).value,0)
         str_deal_closedate = excel_deal_closedate.strftime("%Y-%m-%d")                
               

         # Find corresponding Hubspot Contact VID from actual row 
         # Search for contact via gln 
         company_vid = search_Company(company_gln)
         print("Company Vid:",company_vid)
         # If no contact with the gln exists, Contact needs to be created
         if company_vid == False:
            company_vid=create_Company(company_name, company_street, company_zip, company_city, company_country, company_gln, company_mail, company_phone, company_fax)
         # else go ahead and create deal with this vid 
         else: 
            pass              
         
         # Create dealname from company + unique identifier 
         deal_number = create_DealNumber(company_vid)
         print("Dealnumber:",deal_number)
         dealname = company_name +"- "+ str(deal_number)+ " -"+ str_deal_closedate 

         # Check if the deal still exists hs_vid, deal_date, deal_amount
         deal_exist = search_Deal_Check(amount= deal_betrag, company_name=company_name, closedate=unix_close_date)
         erstzuweisung = check_erstzuweisung(patient_nr=company_patient_nr,deal_closedate=str_deal_closedate) 
         print("Deal exist?:",deal_exist)
         if deal_exist == True:
            status_actual = row - (first_row-1)
            status_last = last_row - (first_row-1)
            print("Deal still exists")
            print("Contact", status_actual,"/",status_last, "Done")

            actual = row - (first_row-1) 
            status_last = last_row -(first_row-1)
            bias = (100/status_last)
            status_actual = round(actual * bias,1)  
            yield "data:" + str(status_actual) + "\n\n"             
            continue

         else: 
            # Create deal in Hubspot via API 
            dealnumber = create_Company_Deal(company_vid,company_zip, company_gln,unix_close_date, deal_gesetz, deal_betrag, dealname,company_patient_nr,erstzuweisung)
            print("Deal created", dealnumber)
            status_actual = row - (first_row-1)
            status_last = last_row - (first_row-1)
            print("Contact", status_actual,"/",status_last, "Done")

            actual = row - (first_row-1) 
            status_last = last_row -(first_row-1)
            bias = (100/status_last)
            status_actual = round(actual * bias,1)  
            yield "data:" + str(status_actual) + "\n\n"       
            
         
         if actual == status_last-1:
               print("delete file")
               os.remove(os.path.join(file_path))
               actual = row - (first_row-1) 
               status_last = last_row -(first_row-1)
               bias = (100/status_last)
               status_actual = round((actual +1) * bias,1)       
               yield "data:" + str(status_actual) + "\n\n"
                           
               break

def test(path):
   print("Path:",path)
   zuweisungen(file_path=path)


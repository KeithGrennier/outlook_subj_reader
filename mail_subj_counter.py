import sys
import win32com.client
import json

def load_json(file_path):
    try:
        with open(file_path, 'r') as in_file:
            return json.load(in_file)
    except json.decoder.JSONDecodeError as e:
        print(f"Error loading JSON from {file_path}: {e}")
        return None

#https://github.com/Hridai/Automating_Outlook/blob/master/ol_script.py
def _find_subfolder(Folders_obj, folder_search_name):
    ''' Recurse through all Outlook folders to find user-defined folder names'''
    for i in range(0, len(Folders_obj)):
        try:
            ret = Folders_obj[i].Folders[folder_search_name]
            return ret
        except:
            ret = _find_subfolder(Folders_obj[i].Folders, folder_search_name)
        if ret is not None:
            return ret
        else:
            continue

def get_emails(keywords:dict,olreadfolder):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = _find_subfolder(outlook.Folders, olreadfolder)
    if inbox is None:
        sys.exit(f'No Folder {olreadfolder} found!!! Exiting.')
    print('Processing Messages')
    messages = inbox.Items
    
    if len(messages) == 0:
        sys.exit('No emails found in folder [{}]'.format(olreadfolder))
    
    mail_counter = 0
    msg_props=None
    print('Reading Mail')
    # print(list(messages)[0])
    for msg in list(messages):
        
        # if mail_counter==0:
        #     msg_props=dir(msg_props)
        #     print(msg_props)
        #     mail_counter+=1
        # listbody = msg.Body.split("\r\n")
        
        for word,count in keywords.items():
            if word.lower() in msg.Subject.lower():
                keywords[word]+=1
        # print(type(msg.Subject))
    return keywords

if __name__ == "__main__":
    word_list=load_json('word_bank.json')
    if word_list is not None:
        word_dict = {word: 0 for word in word_list["words"]}
        subjects_found=get_emails(word_dict,word_list["folder"])
        print(subjects_found)
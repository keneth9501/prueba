"""
This is a detailed example using almost every command of the API
"""
import json
from queue import PriorityQueue
from numpy.core.arrayprint import set_string_function

from numpy.core.numeric import identity
from pandas.core.frame import DataFrame
from pandas.core.indexes.datetimes import date_range
from pandas.io.sql import read_sql_query
from sqlalchemy.sql.expression import null, text, within_group
from report_piki import query
from smtplib import quoteaddr
import time
from sqlalchemy.engine import create_engine, result
import telebot
from telebot import types
from sqlalchemy import text
import pandas as pd
import os
import psycopg2
import time,glob



DATABASES = {
    'piki':{
        'NAME': 'piki',
        'USER': 'power_bi',
        'PASSWORD': 'PowerbI*456$44#',
        'HOST': '162.243.160.180',
        'PORT': 5432,
    },
}

# choose the database to use
db = DATABASES['piki']

# construct an engine connection string
engine_string = "postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}".format(
    user = db['USER'],
    password = db['PASSWORD'],
    host = db['HOST'],
    port = db['PORT'],
    database = db['NAME'],
)

# create sqlalchemy engine
engine = create_engine(engine_string)



query="""SELECT distinct clasificacion  from view_rpt_comercios
where clasificacion not in ('Restaurantes')"""

df=pd.read_sql_query(query,engine)
data=list(df['clasificacion'].values)

# data={
#     "data":df.to_dict(orient="records")
#      }


# response = json.dumps(data,indent=4)

TOKEN = '1917003526:AAEwPVY0aF-V1NIhu0QgI5WVDd-fwoJazVQ'

knownUsers = []  # todo: save these in a file,
userStep = {}  # so they won't reset every time the bot restarts

commands = {  # command description used in the "help" command
    'start'       : 'Get used to the bot',
    'help'        : 'Gives you information about the available commands',
    'sendLongText': 'A test using the \'send_chat_action\' command',
    'getImage'    : 'A test using multi-stage messages, custom keyboard, and media sending',
    'test'        :  'a'

}

hideBoard = types.ReplyKeyboardRemove()


def send_count_comercios(value):
    query='SELECT count(codigo) as data from view_rpt_comercios where clasificacion in ('+"'"+value+"')"
    with engine.connect() as con:
        rs = con.execute(query)
        for row in rs:
         return ( str(row[0]))
   
    
    
    # df=read_sql_query(query,engine)
    # data=list(df['data'].values)
    # string= ' '.join([str(item) for item in data])
 #################################################################
  
# path of the folder
# path = r'C:\Users\50576\Documents\Multipagos\Asignacion_Cartera'
  
# # reading all the excel files
# filenames = glob.glob(path + "\*.xlsx")
# print('File names:', filenames)
  
# # initializing empty data frame
# finalexcelsheet = pd.DataFrame()
  
# # to iterate excel file one by one 
# # inside the folder
# for file in filenames:
  
#     # combining multiple excel worksheets
#     # into single data frames
#     df = pd.concat(pd.read_excel(
#       file, sheet_name=None), ignore_index=True, sort=False)
  
#     # appending excel files one by one
#     finalexcelsheet = finalexcelsheet.append(
#       df, ignore_index=True)
  
# # to print the combined data
# print('Final Sheet:')
  
# finalexcelsheet.to_excel(r'Final.xlsx', index=False)
###########################################################

    
    
    





#error handling if user isn't known yet
#(obsolete once known users are saved to file, because all users
#had to use the /start command and are therefore known to the bot)
def get_user_step(uid):
    if uid in userStep:
        return userStep[uid]
    else:
        knownUsers.append(uid)
        userStep[uid] = 0
        print("New user detected, who hasn't used \"/start\" yet")
        return 0


#only used for console output now
def listener(messages):
    """
    When new messages arrive TeleBot will call this function.
    """
    for m in messages:
        if m.content_type == 'text':
#           print the sent message to the console
            print(str(m.chat.first_name) + " [" + str(m.chat.id) + "]: " + m.text)


bot = telebot.TeleBot(TOKEN)
bot.set_update_listener(listener)  # register listener


#handle the "/start" command
@bot.message_handler(commands=['start'])
def command_start(m):
    cid = m.chat.id
    if cid not in knownUsers:  # if user hasn't used the "/start" command yet:
        knownUsers.append(cid)  # save user id, so you could brodcast messages to all users of this bot later
        userStep[cid] = 0  # save user id and his current "command level", so he can use the "/getImage" command
        bot.send_message(cid, "Hello, stranger, let me scan you...")
        bot.send_message(cid, "Scanning complete, I know you now")
        command_help(m)  # show the new user the help page
    else:
        bot.send_message(cid, "I already know you, no need for me to scan you again!")
@bot.message_handler(commands=['stop'])
def stop(message):
    sent3 = bot.send_message(message.chat.id, 'bye')
    print(sent3)


#help page
@bot.message_handler(commands=['help'])
def command_help(m):
    cid = m.chat.id
    help_text = "The following commands are available: \n"
    for key in commands:  # generate help text out of the commands dictionary defined at the top
        help_text += "/" + key + ": "
        help_text += commands[key] + "\n"
    bot.send_message(cid, help_text)  # send the generated help page


#chat_action example (not a good one...)
@bot.message_handler(commands=['sendLongText'])
def command_long_text(m):
    cid = m.chat.id
    bot.send_message(cid, "If you think so...")
    bot.send_chat_action(cid, 'typing')  # show the bot "typing" (max. 5 secs)
    time.sleep(3)
    bot.send_message(cid, ".")



  #if sent as reply_markup, will hide the keyboard

#user can chose an image (multi-stage command example)
@bot.message_handler(commands=['getImage'])
def command_image(m):
    cid = m.chat.id
    imageSelect = types.ReplyKeyboardMarkup(one_time_keyboard=True,resize_keyboard=True)
    for x in data:
        imageSelect.add(x)
        
    
    bot.send_message(cid, "Please choose your image now", reply_markup=imageSelect)  # show the keyboard
    userStep[cid] = 1  # set the user to the next step (expecting a reply in the listener now)


#if the user has issued the "/getImage" command, process the answer
@bot.message_handler(func=lambda message: get_user_step(message.chat.id) == 1)
def msg_image_select(m):
    cid = m.chat.id
    text = m.text

  # for some reason the 'upload_photo' status isn't quite working (doesn't show at all)
    bot.send_chat_action(cid, 'typing')

    if text != null: 
        bot.send_message(cid, "La cantidad de comerrcios en  \"" + m.text+ "\" es "+send_count_comercios(m.text))
         # send the appropriate image based on the reply to the "/getImage" command
        # bot.send_photo(cid, open('rooster.jpg', 'rb'),
        #                reply_markup=hideBoard) 
    #                 send file and hide keyboard, after image is sent
        userStep[cid] = 1  # reset the users step back to 0
    # elif text == 'Bebidas':
    #     bot.send_document(cid, open('pandas_multiple.xlsx', 'rb'), reply_markup=hideBoard)
        
    #     userStep[cid] = 0
    else:
        bot.send_message(cid, "Please, use the predefined keyboard!")
        bot.send_message(cid, "Please try again")
        userStep[cid] = 0


#filter on a specific message
@bot.message_handler(func=lambda message: message.text == "hi")
def command_text_hi(m):
    bot.send_message(m.chat.id, "I love you too!")

#filter on a database


#default handler for every other text
@bot.message_handler(func=lambda message: True, content_types=['text'])
def command_default(m):
#    this is the standard reply to a normal message
    bot.send_message(m.chat.id, "I don't understand \"" + m.text+ "\"\nMaybe try the help page at /help")
    


bot.polling()

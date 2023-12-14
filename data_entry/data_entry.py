from pathlib import Path
import random
import PySimpleGUI as sg
import pandas as pd
import datetime


# Add some color to the window
sg.theme('DarkTeal9')

current_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXCEL_FILE = current_dir / 'Data_Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)



SYMBOL_RIGHT =    '▶'
SYMBOL_DOWN =  '▼'

def collapse(layout, key):

    # sg.pin allows us to diplay or hide the column
    return sg.pin(sg.Column(layout, key=key))
section_1 = [
            [sg.Text('Elimu ya Mabadiliko',font=('Montserat',13,'bold'))], 
            [sg.Text('Majadaliano ya tabia hatarishi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Majadaliano ya tabia hatarishi zinazopelekea kupata maambukizi ya VVU', default_value='Yes'),
                sg.Text('Majadiliano ya matumizi sahihi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Majadiliano ya matumizi sahihi na endelevu ya Kondomu', default_value='Yes'),
                sg.Text('Elimu ya ukatili wa Kijinsia', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya ukatili wa Kijinsia', default_value='Yes'),
                sg.Text('Elimu ya uzazi wa mpango', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya uzazi wa Mpango', default_value='Yes'),
                sg.Text('Kuwashirikisha wengine hali yako ya maambukizi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kuwashirikisha wengine hali yako ya maambukizi.1', default_value='Yes')],
            
            [sg.Text('Ufuasi wa matibabu ya VVU', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ufuasi wa matibabu ya VVU', default_value='Yes'),
                sg.Text('Elimu ya makuzi na malezi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya makuzi na malezi', default_value='Yes'),
                sg.Text('Elimu ya upimaji wa VVU', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya upimaji wa VVU', default_value='Yes'),
                sg.Text('Elimu ya simu bila malipo 117', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya simu bila malipo 117', default_value='Yes'),
                sg.Text('Elimu ya ujumbe bila malipo (15017) ', size=(15,1)), sg.Combo(['Yes', 'No'], key='Elimu ya ujumbe bila malipo (15017) ', default_value='Yes')],

            [sg.Text('HUDUMA ZA KITABIBU',font=('Montserat',13,'bold'))],    

            [sg.Text('Huduma ya upimaji wa VVU', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya upimaji wa VVU', default_value='No'),
                sg.Text('Huduma ya Kliniki ya Kifua kikuu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya Kliniki ya Kifua kikuu', default_value='No'),
                sg.Text('Huduma ya uzazi wa mpango', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya uzazi wa mpango', default_value='No'),
                sg.Text('Ugawaji wa Kondomu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ugawaji wa Kondomu', default_value='Yes'),
                sg.Text('Kliniki ya Kinga ya VVU', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kliniki ya Kinga ya VVU  toka kwa mama mjamzito kwenda kwa mtoto (PMTCT)', default_value='No')],
            
            [sg.Text('Kliniki ya magonjwa ya ngono', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kliniki ya magonjwa ya ngono', default_value='No'),
                sg.Text('Kliniki ya matibabu ya VVU', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kliniki ya matibabu ya VVU', default_value='No'),
                sg.Text('Matibabu ya dawa za kulevya', size=(15,1)), sg.Combo(['Yes', 'No'], key='Matibabu ya dawa za kulevya', default_value='No'),
                sg.Text('Huduma za saratani ya shingo ya kizazi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma za saratani ya shingo ya kizazi', default_value='No')],

            [sg.Text('HUDUMA ZA MTAMBUKO',font=('Montserat',13,'bold'))],    

            [sg.Text('Ugawaji wa Taulo za kike', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ugawaji wa Taulo za kike', default_value='No'),
                sg.Text('Kusaidia kuibua', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kusaidia kuibua na  kutatua tatizo la Ukatili wa kijinsia', default_value='No'),
                sg.Text('Malipo ya biashara (Seed money)', size=(15,1)), sg.Combo(['Yes', 'No'], key='Malipo ya biashara (Seed money)', default_value='No'),
                sg.Text('Huduma ya unyanyapaa', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya unyanyapaa na kutengwa katika jamii', default_value='No'),
                sg.Text('Huduma ya 117', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya simu bila malipo 117', default_value='Yes')],
            
            [sg.Text('Huduma ya  (15017)', size=(15,1)), sg.Combo(['Yes', 'No'], key='Huduma ya ujumbe bila malipo (15017)', default_value='Yes'),
                sg.Text('Rufaa kwa ajili ya huduma za ukatili', size=(15,1)), sg.Combo(['Yes', 'No'], key='Rufaa kwa ajili ya huduma za ukatili wa kijinsia (Polisi/ kituo cha Afya/Afisa Ustawi wa jamii/Serikali ya kijiji/msaada wa kisheria/n.k', default_value='No'),
                sg.Text('Rufaa kwa ajili ya huduma ya kiuchumi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Rufaa kwa ajili ya huduma ya kiuchumi ( ofisi ya maendeleo ya jamii/kikundi cha uzalishaji mali/benki/n.k.', default_value='No'),
                sg.Text('Rufaa nyinginezo', size=(15,1)), sg.Combo(['Yes', 'No'], key='Rufaa nyinginezo  (Taja)……………..', default_value='No')],

            [sg.Text('VIGEZO HATARISHI',font=('Montserat',13,'bold'))],    

            [sg.Text('Ameacha shule', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ameacha shule', default_value='Yes'),
                sg.Text('Kiongozi wa kaya', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kiongozi wa kaya', default_value='No'),
                sg.Text('Yatima (wote wawili)', size=(15,1)), sg.Combo(['Yes', 'No'], key='Yatima (Asiye na wazazi wote wawili)', default_value='No'),
                sg.Text('Ameolewa katika umri mdogo', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ameolewa katika umri mdogo', default_value='No'),
                sg.Text('Amezaa katika umri mdogo', size=(15,1)), sg.Combo(['Yes', 'No'], key='Amezaa katika umri mdogo', default_value='No')],
            
            [sg.Text('Ni mjamzito katika umri mdogo', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ni mjamzito katika umri mdogo (chini ya miaka 18)', default_value='No'),
                sg.Text('Amewahi kupata magonjwa ya ngono', size=(15,1)), sg.Combo(['Yes', 'No'], key='Amewahi kupata magonjwa ya ngono', default_value='No'),
                sg.Text('Anamahusiano ya kimapenzi na wanaume wengi/kumzidi umri wa miaka 5 au zaidi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Anamahusiano ya kimapenzi na wanaume wengi/kumzidi umri wa miaka 5 au zaidi', default_value='No'),
                sg.Text('Anafanya ngono kwa ajili ya kujipatia kipato au zawadi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Anafanya ngono kwa ajili ya kujipatia kipato au zawadi', default_value='No')],

            [sg.Text('hajawahi tumia kondomu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Hana uelewa wa matumizi sahihi ya kondomu/hajawahi tumia kondomu', default_value='Yes'),
                sg.Text('Ni Mlevi/anatumia madawa ya kulevya', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ni Mlevi/anatumia madawa ya kulevya', default_value='No'),
                sg.Text('Amefanyiwa ukatili wa kijinsia', size=(15,1)), sg.Combo(['Yes', 'No'], key='Amefanyiwa ukatili wa kijinsia', default_value='No'),
                sg.Text('Mlemavu asiejimudu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Mlemavu asiejimudu', default_value='No')],
                
            [sg.Text('Mtoto wa mtaani asie na wazazi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Mtoto wa mtaani asie na wazazi', default_value='No'),
                sg.Text('Ana mtoto kwa sasa', size=(15,1)), sg.Combo(['Yes', 'No'], key='Ana mtoto kwa sasa', default_value='No'),
                sg.Text('Anatokea kwenye kaya maskini (TASAF)', size=(15,1)), sg.Combo(['Yes', 'No'], key='Tasaf', default_value='Yes')],
        ]

section_3 = [
    [sg.Text('Umuhimu wa elimu ya malezi katika familia/jamii', size=(15,1)), sg.Combo(['Yes', 'No'], key='Umuhimu wa elimu ya malezi katika familia/jamii', default_value='Yes'),
        sg.Text('Maana ya ukatili dhidi ya watoto na aina za ukatili', size=(15,1)), sg.Combo(['Yes', 'No'], key='Maana ya ukatili dhidi ya watoto na aina za ukatili', default_value='Yes'),
        sg.Text('Visababishi na mambo hatarishi yanayochangia ukatili dhidi ya watoto', size=(15,1)), sg.Combo(['Yes', 'No'], key='Visababishi na mambo hatarishi yanayochangia ukatili dhidi ya watoto', default_value='Yes'),
        sg.Text('Athari za ukatili dhidi ya watoto katika familia/jamii zetu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Athari za ukatili dhidi ya watoto katika familia/jamii zetu', default_value='Yes'),
        sg.Text('Malezi ya watoto wadogo yanayozingatia ulinzi na usalama', size=(15,1)), sg.Combo(['Yes', 'No'], key='Malezi ya watoto wadogo yanayozingatia ulinzi na usalama', default_value='Yes')],
    
    [sg.Text('Kujenga na kudumisha mazingira salama ya watoto nyumbani na katika jamii', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kujenga na kudumisha mazingira salama ya watoto nyumbani na katika jamii', default_value='Yes'),
        sg.Text('Stadi muhimu zinazofundisha watoto kujilinda na ukatili', size=(15,1)), sg.Combo(['Yes', 'No'], key='Stadi muhimu zinazofundisha watoto kujilinda na ukatili', default_value='Yes'),
        sg.Text('Mawasiliano yanayofaa kwa watoto wa rika tofauti', size=(15,1)), sg.Combo(['Yes', 'No'], key='Mawasiliano yanayofaa kwa watoto wa rika tofauti', default_value='Yes'),
        sg.Text('Kuwawezesha watoto kuzungumza na utaratibu wa rufaa katika jamii', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kuwawezesha watoto kuzungumza na utaratibu wa rufaa katika jamii', default_value='Yes'),
        sg.Text('Haki za mtoto na sheria inayomlinda mtoto katika muktadha wa malezi', size=(15,1)), sg.Combo(['Yes', 'No'], key='Haki za mtoto na sheria inayomlinda mtoto katika muktadha wa malezi', default_value='Yes')],
    
    [sg.Text('Malezi ya watoto wenye ulemavu', size=(15,1)), sg.Combo(['Yes', 'No'], key='Malezi ya watoto wenye ulemavu', default_value='Yes'),
        sg.Text('Jinsia, Mila na Desturi katika Malezi ya watoto', size=(15,1)), sg.Combo(['Yes', 'No'], key='Jinsia, Mila na Desturi katika Malezi ya watoto', default_value='Yes'),
        sg.Text('Kuimarisha uchumi wa kaya', size=(15,1)), sg.Combo(['Yes', 'No'], key='Kuimarisha uchumi wa kaya', default_value='Yes'),
        sg.Text('Utamaduni na vyombo vya habari', size=(15,1)), sg.Combo(['Yes', 'No'], key='Utamaduni na vyombo vya habari', default_value='Yes')],

]
section_2 = [
    [sg.Text('WATU WAKARIBU' ,font=('Montserat',13,'bold'))],
    [sg.Text('Namba ya simu', size=(10,1)), sg.InputText(key='Namba ya simu'),
        sg.Text('Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', size=(10,1)), sg.Combo(['Mzazi','Mlezi','Rafiki','Mwenza','Jirani'], key='Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', default_value='Rafiki'),
        sg.Text('Umri', size=(15,1)), sg.Spin([i for i in range(10,50)],
                                                        initial_value=0, key='Umri'),
        sg.Text('Jinsi(ME/KE)', size=(15,1)), sg.Combo(['Me','Ke'], key='Jinsi(ME/KE)', default_value='Ke'),],
    # [sg.Text('Namba ya simu', size=(10,1)), sg.InputText(key='Namba ya simu 2'),
    #     sg.Text('Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', size=(10,1)), sg.Combo(['Mzazi','Mlezi','Rafiki','Mwenza','Jirani'], key='Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza 2', default_value='Rafiki'),
    #     sg.Text('Umri', size=(15,1)), sg.Spin([i for i in range(10,50)],
    #                                                     initial_value=0, key='Umri2'),
    #     sg.Text('Jinsi(ME/KE)', size=(15,1)), sg.Combo(['Me','Ke'], key='Jinsi(ME/KE)2', default_value='Ke'),],
    # [sg.Text('Namba ya simu', size=(10,1)), sg.InputText(key='Namba ya simu3'),
    #     sg.Text('Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', size=(10,1)), sg.Combo(['Mzazi','Mlezi','Rafiki','Mwenza','Jirani'], key='Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza 3', default_value='Rafiki'),
    #     sg.Text('Umri', size=(15,1)), sg.Spin([i for i in range(10,50)],
    #                                                     initial_value=0, key='Umri3'),
    #     sg.Text('Jinsi(ME/KE)', size=(15,1)), sg.Combo(['Me','Ke'], key='Jinsi(ME/KE)3', default_value='Ke'),],
    # [sg.Text('Namba ya simu', size=(10,1)), sg.InputText(key='Namba ya simu4'),
    #     sg.Text('Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', size=(10,1)), sg.Combo(['Mzazi','Mlezi','Rafiki','Mwenza','Jirani'], key='Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza 4', default_value='Rafiki'),
    #     sg.Text('Umri', size=(15,1)), sg.Spin([i for i in range(10,50)],
    #                                                     initial_value=0, key='Umri4'),
    #     sg.Text('Jinsi(ME/KE)', size=(15,1)), sg.Combo(['Me','Ke'], key='Jinsi(ME/KE)4', default_value='Ke'),],
    # [sg.Text('Namba ya simu', size=(10,1)), sg.InputText(key='Namba ya simu5'),
    #    sg.Text('Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza', size=(10,1)), sg.Combo(['Mzazi','Mlezi','Rafiki','Mwenza','Jirani'], key='Uhusiano (Mzazi/Mlezi/Rafiki/Mwenza 5', default_value='Rafiki'),
    #     sg.Text('Umri', size=(15,1)), sg.Spin([i for i in range(10,50)],
    #                                                     initial_value=0, key='Umri5'),
    #     sg.Text('Jinsi(ME/KE)', size=(15,1)), sg.Combo(['Me','Ke'], key='Jinsi(ME/KE)5', default_value='Ke'),],
]
def validate_date(a,b):
    try:
        datetime.datetime.strptime(a, '%d/%m/%Y')
        datetime.datetime.strptime(b, '%d/%m/%Y')
        return True, ""
    except ValueError:
        return False, "Invalid date format. Use DD/MM/YYYY."

def validate_phone_number(phone_number):
    if len(phone_number) != 10:
        return False
    for char in phone_number:
        if not char.isdigit():
            return False
    return True

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Taarifa za muelimishaji',font=('Montserat',13,'bold'))],
    [sg.Text('Tarehe ya kutoa hudma', size=(10,1)), sg.InputText(key='Tarehe ya kutoa hudma',enable_events=True,)],
        # sg.CalendarButton("DOB",format='%d/%m/%Y', close_when_date_chosen=True, target='Tarehe ya kutoa hudma', no_titlebar=False)],
    
    [sg.Text(key='output', size=(40, 1),text_color='red')],
    [sg.Text('Name', size=(10,1)), sg.InputText(key='Jina la kwanza la muelimishaji rika'),
        sg.Text('SName', size=(10,1)), sg.InputText(key='Jina la pili la muelimishaji rika'),
        sg.Text('LName', size=(10,1)), sg.InputText(key='Jina la tatu la muelimishaji rika')],

    [sg.Text('Phone', size=(10,1)), sg.InputText(key='Namba ya simu ya muelimishaji rika', enable_events=True)],
    [sg.Text(key='error_msg',text_color='red')],
    [sg.Text('Kijiji', size=(10,1)), sg.InputText(key='Kijiji'),
        sg.Text('Kata', size=(10,1)), sg.InputText(key='Kata'),
        sg.Text('Wilaya', size=(10,1)), sg.InputText(key='Wilaya')],

    [sg.Text('Mkoa', size=(10,1)), sg.InputText(key='Mkoa')],

    [sg.Text('Taarifa za mteja',font=('Montserat',13,'bold'))],
    #[sg.Text('kijiji', size=(10,1)), sg.InputText(key='Kijiji1')],
    [sg.Text('CName', size=(10,1)), sg.InputText(key='Jina la kwanza la mteja'),
        sg.Text('CSName', size=(10,1)), sg.InputText(key='Jina la pili la mteja'),
        sg.Text('CLName', size=(10,1)), sg.InputText(key='Jina la tatu la mteja')],

    [sg.Text('ID NUMBER', size=(10,1)), sg.InputText(key='Namba ya utambulisho ya mteja')],

    [sg.Text('Tarehe ya kuzaliwa ya mteja', size=(10,1)), sg.InputText(key='Tarehe ya kuzaliwa ya mteja',enable_events=True,),
        # sg.CalendarButton("DOB",format='%d/%m/%Y', close_when_date_chosen=True, target='Tarehe ya kuzaliwa ya mteja', no_titlebar=False),
        sg.Text('Namba ya simu ya mteja', size=(10,1)), sg.InputText(key='Namba ya simu ya mteja', enable_events=True),
        sg.Text('Kazi', size=(15,1)), sg.Combo(['Mkulima', 'Mjasiriamali','Muajiriwa','Mwnafunzi','Hanakazi'], key='Kuwashirikisha wengine hali yako ya maambukizi', default_value='Mkulima')],
        [sg.Text(key='output2', size=(40, 1),text_color='red')],

    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC1-'),
    sg.Text('HUDUMA MABAMBALI ALIZOTOA MUELIMISHSAJIRIKA KWA MTEJA WAKE KWA MWEZI MZIMA',font=('Montserat',13,'bold')),],
    [collapse(section_1, '-SEC_1-')],

    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC2-'),
    sg.Text('WATU WA WAKARIBU',font=('Montserat',13,'bold'))],
    [collapse(section_2, '-SEC_2-')],

    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC3-'),
    sg.Text('MADA ALIZO FUNDISHA MAMA KIJAN KWA WENZAKE',font=('Montserat',13,'bold'))],
    [collapse(section_3, '-SEC_3-')],

    


    [sg.Submit(), sg.Button('Clear'), sg.Exit(button_color=('white', 'red'))]
    
]

window = sg.Window('Simple data entry form', layout)

#------------------------------------------------------------------
start_date = datetime.date(1997, 1, 1)
end_date = datetime.date(2005, 12, 31)



days_between = (end_date - start_date).days

#------------------------------------------------------------------
# Load the excel file into a pandas dataframe
# df = pd.read_excel("agy.xlsx")

# Print the entire dataframe
# print(df)

# # Access specific columns of the dataframe
# column_1 = df['First Name'].tolist()
# column_2 = df['Middle Name'].tolist()
# column_3 = df['Last Name'].tolist()
# column_4 = df['Phone'].tolist()

# column_First = column_1
# column_Middle = column_2
# column_Last = column_3
# column_Phone = column_4

#--------------to generate phone number---------------------
phonecodes = ['74', '75', '76','67','68', '69', '77','79','64', '66', '78','71','72','62']

def generate_tanzanian_phone_number():
    network_code = ''.join([str(random.randint(3, 9)) for i in range(1)])
    subscriber_number = ''.join([str(random.randint(0, 9)) for i in range(6)])
    network = random.choice(phonecodes)
    return f"0{network}{network_code}{subscriber_number[:3]}{subscriber_number[3:]}"

#--------------to generate phone number---------------------
phone_numbers = []
for i in range(2000):
    numbers = generate_tanzanian_phone_number()
    phone_numbers.append(numbers)

def clear_input():
    for key in values:
        window[key]('')
    return None
count = 0
opened1 = True
opened2 = True
opened3 = True
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        confirm_exit = sg.popup_yes_no("Do you realy want to Exit")
        if confirm_exit == 'Yes':
            break
    # if event == 'Clear':
    #     clear_input()
    if event == '-OPEN_SEC1-':
        opened1 = not opened1
        window['-OPEN_SEC1-'].update(SYMBOL_DOWN if opened1 else SYMBOL_RIGHT)
        window['-SEC_1-'].update(visible=opened1)
    
    if event == '-OPEN_SEC2-':
        opened2 = not opened2
        window['-OPEN_SEC2-'].update(SYMBOL_DOWN if opened2 else SYMBOL_RIGHT)
        window['-SEC_2-'].update(visible=opened2)

    if event == '-OPEN_SEC3-':
        opened3 = not opened3
        window['-OPEN_SEC3-'].update(SYMBOL_DOWN if opened3 else SYMBOL_RIGHT)
        window['-SEC_3-'].update(visible=opened3)

    try:
        if event == 'Submit':
            #---------------create phone ---------------------
            selected_contact = random.choice(phone_numbers)
            selected_contact2 = random.choice(phone_numbers)
            selected_contact3 = random.choice(phone_numbers)
            selected_contact4 = random.choice(phone_numbers)
            selected_contact5 = random.choice(phone_numbers)
            #---------------names---------------------
            # column_First = random.choice(column_First)
            # column_Middle = random.choice(column_Middle)
            # column_Last = random.choice(column_Last)
            # column_Phone = random.choice(column_Phone)

            # --------------date------------------
            # random_num_days = random.randint(0, days_between)
            # random_date = start_date + datetime.timedelta(days=random_num_days)
            # formatted_date = random_date.strftime("%d/%m/%Y")



          #----------calculate age ------------------------
            def calculate_age(birth_date):
                today = datetime.datetime.strptime(values["Tarehe ya kutoa hudma"], "%d/%m/%Y").date()
                age = today.year - birth_date.year
                return age

            birth_date = datetime.datetime.strptime(values["Tarehe ya kuzaliwa ya mteja"], "%d/%m/%Y").date()
            age = calculate_age(birth_date)
            #----------------counter submision------------------
            if count < 10000:
                count += 1


            date_string = values["Tarehe ya kutoa hudma"]

            # Split the date string into day, month, and year
            day, month, year = date_string.split('/')

            # Convert the components to integers if needed
            day = int(day)
            month = int(month)
            year = int(year)
            values['Tarehe'] = day
            values['Mwezi'] = month
            values['Mwaka'] = year
            values['robo'] = "Q11"
            values["Tarehe ya kuzaliwa ya mteja"]
            values["Jina la kwanza la muelimishaji rika"] = values["Jina la kwanza la muelimishaji rika"].title()
            values["Jina la pili la muelimishaji rika"] = values["Jina la pili la muelimishaji rika"].title()
            values["Jina la tatu la muelimishaji rika"] = values["Jina la tatu la muelimishaji rika"].title()
            values["Kijiji"] = values["Kijiji"].title()
            values["Kata"] = values["Kata"].title()
            values["Wilaya"] = values["Wilaya"]
            values["Mkoa"] = values["Mkoa"].title()
            values["Kijiji1"] = values["Kijiji"].title()
            values["Jina la kwanza la mteja"] = values["Jina la kwanza la mteja"].title()
            values["Jina la pili la mteja"] = values["Jina la pili la mteja"].title()
            values["Jina la tatu la mteja"] = values["Jina la tatu la mteja"].title()
            values["Umri wa mteja"] = age
            values["Namba ya simu ya mteja"] = values["Namba ya simu ya mteja"]
            values["Namba ya simu"] = values["Namba ya simu"]
            # values["Namba ya simu3"] = int(values["Namba ya simu3"])
            # values["Namba ya simu4"] = int(values["Namba ya simu4"])
            # values["Namba ya simu5"] = int(values["Namba ya simu5"])

            # values["Namba ya simu 2"] = values["Namba ya simu 2"]
            # values["Namba ya simu3"] = selected_contact3
            # values["Namba ya simu4"] = selected_contact4
            # values["Namba ya simu5"] = selected_contact5

            new_record = pd.DataFrame(values, index=[0])
            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup_timed('Data saved!, counter:'+ str(count),auto_close_duration=3)

    except ValueError:
        sg.popup_error('Kunashid kwenye tarehe')
    
    except TypeError:
        sg.popup_error('UMEKOSEA TAREHE')
    
     #values["Namba ya simu3"] = selected_contact3
        #values["Namba ya simu4"] = selected_contact4
        #values["Namba ya simu5"] = selected_contact5
        
        # try:
        #     values['Tarehe ya kuzaliwa ya mteja']
        # except ValueError:
        #     sg.popup('Wrong mteja Date Format')
        # try:
        #     values['Tarehe ya kutoa hudma']
        # except ValueError:
        #     sg.popup('Wrong huduma Date Format')
        # else:
        #     new_record = pd.DataFrame(values, index=[0])
        #     df = pd.concat([df, new_record], ignore_index=True)
        #     df.to_excel(EXCEL_FILE, index=False)
        #     sg.popup('Data saved!, counter:'+ str(count))
        
    is_valid, error_msg = validate_date(values['Tarehe ya kutoa hudma'],values['Tarehe ya kuzaliwa ya mteja'],)
    if not is_valid:
        window['output'].update(error_msg)
    else:
        window['output'].update("")
    if not is_valid:
        window['output2'].update(error_msg)
    else:
        window['output2'].update("")
    phone_number = values['Namba ya simu ya muelimishaji rika']
    if not validate_phone_number(phone_number):
        window['error_msg'].update('Invalid phone number')
    else:
        window['error_msg'].update('')
    
    
    
window.close()

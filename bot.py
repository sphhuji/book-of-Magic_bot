import discord
import asyncio
from discord.ext import commands
import random
import openpyxl
import time
import os

bot = commands.Bot(command_prefix='+)!*&)')

@bot.event
async def on_ready():
    print(f"We have logged in as {bot.user}")
    await bot.change_presence(activity=discord.Activity(type=discord.ActivityType.watching, name =f"=help"))
    print(discord.__version__)

@bot.event
async def on_message(message):
    if message.content.startswith('=help'):
        embed=discord.Embed(title="마법의 책 사용법", description="너의 지식을 추가해주려면\n=작성 [단어] [뜻]\n(지식은 255개까지밖에 저장할 수 없고 띄어쓰기는 인식하지 못해!)\n\n책을 읽으려면\n=독서 [단어]\n\n무언가를 소환하고 싶으면 =소환 을 외쳐봐!\n\n새로 추가된 기능을 보고 싶으면 =패치노트 주문을 외워!", color=0x62c1cc)

        await message.channel.send(embed=embed)

    if message.content.startswith('=크시'):
        await message.channel.send('음...그래, 뭐')

    if message.content.startswith('=제작자'):
        await message.channel.send('허지#5733')

    if message.content.startswith('=모렐로노미콘'):
        await message.channel.send('뭐, 좋아 좋지.')

    if message.content.startswith('=공백'):
        await message.channel.send('ㅤ')

    if message.content.startswith('=패치노트'):
        embed = discord.Embed(title="패치노트", description="1. =패치노트 기능 추가\n\n2. =소환 기능 추가", color=0x62c1cc)
        embed.set_footer(text="패치일 4월 1일")

        await message.channel.send(embed=embed)

    if message.content.startswith('=소환'):
        embed = discord.Embed(title="소환!!", description="(ノ°∀°)ノ⌒･*:.｡. .｡.:*･゜ﾟ･*☆", color=0x800080)

        await message.channel.send(embed=embed)
        time.sleep(3)
        
        summon = ["소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/%eb%ac%b4%ec%95%bc%ed%98%b8-%eb%ac%b4%ed%95%9c%eb%8f%84%ec%a0%84-%ec%95%b5%ec%bb%a4%eb%a6%ac%ec%a7%80-%ed%95%9c%ec%9d%b8%ed%9a%8c%ea%b4%80-%ea%b7%b8%eb%a7%8c%ed%81%bc-gif-20464026", "소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/%eb%a9%88%ec%b6%b0-gif-20730775", "소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/dancing-coffin-coffin-dance-funeral-funny-farewell-gif-16737844", "소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/nick-young-question-mark-huh-what-confused-gif-4995479", "소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/crying-black-guy-meme-sad-gif-11746329", "소환성공! (*・ω・)ﾉ\nhttps://tenor.com/view/%ed%83%9c%eb%b3%b4%ed%95%b4-taebo-diet-exercise-getting-fit-gif-14365603 ", "소환실패...(ﾉД`)"]

        randomNum = random.randrange(0, len(summon))
        print("랜덤수 값 :" + str(randomNum))
        print(summon[randomNum])
        await message.channel.send(summon[randomNum])

    if message.content.startswith('=작성'):
        file = openpyxl.load_workbook("book-of-Magic_bot\지식.xlsx")
        sheet = file.active
        learn = message.content.split(" ")
        for i in range(1, 255):
            if sheet["A" + str(i)].value == "-" or sheet["A" + str(i)].value == learn[1]:
                sheet["A" + str(i)].value = learn[1]
                sheet["B" + str(i)].value = learn[2]
                break
        file.save("book-of-Magic_bot\지식.xlsx")
        await message.channel.send('작성 완료!')

    if message.content.startswith('=독서'):
        file = openpyxl.load_workbook("book-of-Magic_bot\지식.xlsx")
        sheet = file.active
        memory = message.content.split(" ")
        for i in range(1, 255):
           if sheet["A" + str(i)].value == memory[1]:
               await message.channel.send(sheet["B" + str(i)].value)
               break

bot.run(os.environ['token'])

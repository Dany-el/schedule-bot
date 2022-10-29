import logging
import xlrd
import datetime

from aiogram import Bot, Dispatcher, executor, types

API_TOKEN = '5736608244:AAFLxV9ChfUjnNow75b2IVR0UpXOgL8GQxs'

# Configure logging
logging.basicConfig(level=logging.INFO)

# Initialize bot and dispatcher
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

# Week number
my_date = datetime.date.today()
year, week_num, day_of_week = my_date.isocalendar()

is_have_pair_unpair = False


@dp.message_handler(commands='start')
async def welcome(message: types.Message):
    await message.reply(f"Hello, {message.from_user.first_name} {message.from_user.last_name} !")

    keyboard_markup = types.InlineKeyboardMarkup(row_width=2)

    text_and_data = (
        ('Yes', 'yes'),
        ('No', 'no'),
    )

    row_btns = (types.InlineKeyboardButton(text, callback_data=data) for text, data in text_and_data)
    keyboard_markup.row(*row_btns)

    await bot.send_message(message.chat.id,
                           "Do you have pair and unpair weeks?", reply_markup=keyboard_markup)


@dp.callback_query_handler(text='no')  # if cb.data == 'no'
@dp.callback_query_handler(text='yes')  # if cb.data == 'yes'
async def inline_kb_answer_callback_handler(query: types.CallbackQuery):
    answer_data = query.data
    global is_have_pair_unpair
    # await query.answer(f'You answered with {answer_data!r}')

    if answer_data == 'yes':
        is_have_pair_unpair = True
    elif answer_data == 'no':
        is_have_pair_unpair = False
    await bot.send_message(query.from_user.id, "Good! Let's start!")


@dp.message_handler(commands=['help'])
async def send_welcome(message: types.Message):
    await message.reply("*Available commands:"
                        "\n-------------------------------------"
                        "\n/help - show available commands"
                        "\n/day  - show current day schedule"
                        "\n/week - show current week schedule"
                        "\nmessage 'pair' to show picture"
                        "\nmessage 'unpair' to show picture"
                        "\n-------------------------------------*", parse_mode='Markdown')


workbook = xlrd.open_workbook("venv/schedule/AC-227.xlsx")


@dp.message_handler(commands=['week'])
async def send_week(message: types.Message):
    global workbook
    # Week number is pair or not
    if is_have_pair_unpair:
        if week_num % 2 == 0:
            worksheet = workbook.sheet_by_index(0)
        else:
            worksheet = workbook.sheet_by_index(1)
    else:
        worksheet = workbook.sheet_by_index(0)

    schedule = ""

    # Iterate the rows and columns
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            schedule += worksheet.cell_value(i, j) + "  "
        schedule += '\n'
    await message.reply(f"*{schedule}*", disable_web_page_preview=True, parse_mode='Markdown')


@dp.message_handler(commands=['day'])
async def send_day(message: types.Message):
    day_name = my_date.strftime('%A')
    global workbook
    # Week number is pair or not
    if is_have_pair_unpair:
        if week_num % 2 == 0:
            worksheet = workbook.sheet_by_index(0)
        else:
            worksheet = workbook.sheet_by_index(1)
    else:
        worksheet = workbook.sheet_by_index(0)

    schedule = ""

    forward = False

    if day_name != "Sunday" or day_name != "Saturday":
        # Iterate the rows and columns
        for i in range(0, worksheet.nrows):
            for j in range(0, worksheet.ncols):
                if forward:
                    schedule += worksheet.cell_value(i, j) + "  "
                if worksheet.cell_value(i, j) == day_name and worksheet.cell_value(i, j + 1) == "":
                    schedule += worksheet.cell_value(i, j) + "'s schedule\n"
                    forward = True
                    break
                elif worksheet.cell_value(i, j) == "":
                    forward = False
            if forward:
                schedule += '\n'
        await message.reply(f"*{schedule}*", disable_web_page_preview=True, parse_mode='Markdown')
    else:
        await message.reply("*You have a rest today!*", parse_mode='Markdown')


@dp.message_handler(regexp='^[pP][aA][Ii][Rr]')
async def pair_week(message: types.Message):
    with open('venv/img/Pair.png', 'rb') as photo:
        await message.reply_photo(photo, caption='Pair week')


@dp.message_handler(regexp='^[Uu][Nn][pP][aA][Ii][Rr]')
async def unpair_week(message: types.Message):
    with open('venv/img/Unpair.png', 'rb') as photo:
        await message.reply_photo(photo, caption='Unpair week')


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)

from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup

button_1 = InlineKeyboardButton(text="Создать файл", callback_data="restart")

kb = InlineKeyboardMarkup(inline_keyboard=[[button_1]])

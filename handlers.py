from aiogram import Router, F, types
from aiogram.filters import StateFilter, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State, default_state
from aiogram.types import Message, CallbackQuery

from bot import bot
from excel_util import open_exel_1, open_exel_2, create_excel
from keyboard import kb

router =Router()

class FSMFillForm(StatesGroup):
    send_excel_1 = State()
    send_excel_2 = State()

@router.message(CommandStart(), StateFilter(default_state))
async def process_start(mes: Message, state: FSMContext):
    await mes.answer(text="Добрый день. Скиньте первый файл excel")
    await state.set_state(FSMFillForm.send_excel_1)

@router.callback_query(F.data == "restart", StateFilter(default_state))
async def restart(cb: CallbackQuery, state: FSMContext):
    await bot.send_message(
        chat_id=cb.from_user.id,
        text="Добрый день. Скиньте первый файл excel")
    await state.set_state(FSMFillForm.send_excel_1)


@router.message(F.document, StateFilter(FSMFillForm.send_excel_1))
async def send_excel_1(message: Message, state: FSMContext):
    await bot.download(
        message.document,
        destination=f"excel_1_{message.from_user.id}.xlsx"
    )
    await state.set_state(FSMFillForm.send_excel_2)
    await message.answer("Скиньте второй файл excel")


@router.message(F.document, StateFilter(FSMFillForm.send_excel_2))
async def send_excel_2(message: Message, state: FSMContext, mes=None):
    await bot.download(
        message.document,
        destination=f"excel_2_{message.from_user.id}.xlsx"
    )
    try:
        dct = open_exel_1(f"excel_1_{message.from_user.id}.xlsx")
        res = open_exel_2(f"excel_2_{message.from_user.id}.xlsx", dct)
        create_excel(res, f"excel_3_{message.from_user.id}.xlsx")
        await message.answer_document(
            types.FSInputFile(path=f"excel_3_{message.from_user.id}.xlsx"),
            caption="Вот ваш файл, для начала загрузки новых файлов нажмите кнопку",
            reply_markup=kb
        )
    except Exception as e:
        await message.answer(text=f"Видимо вы скинули не те файлы excel, произошла ошибка -\n "
                                  f"{e}\nНажмите кнопку чтобы начать заново", reply_markup=kb)
    await state.set_state(default_state)


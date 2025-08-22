import os
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, executor, types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from apscheduler.schedulers.asyncio import AsyncIOScheduler
import openpyxl

API_TOKEN = "8052107359:AAE7F0Dui6lWw_ejUvx0PR1GaJBzvu9b-t0"  # встав свій токен
ADMIN_ID = 1418044149              # встав свій Telegram ID

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

# Excel-файл
FILE_NAME = "zayavky.xlsx"

# Створення Excel при старті
def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Заявки"
        ws.append(["Дата", "ПІБ", "Адреса", "Телефон", "Проблема"])
        wb.save(FILE_NAME)

init_excel()

# Стан машини для заявки
class Form(StatesGroup):
    fio = State()
    address = State()
    phone = State()
    problem = State()

# Меню
main_menu = types.ReplyKeyboardMarkup(resize_keyboard=True)
main_menu.add("📝 Залишити заявку")

@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    await message.answer("Ласкаво просимо до ТРК «Мир ТВ» Покров", reply_markup=main_menu)

# Початок заявки
@dp.message_handler(lambda message: message.text == "📝 Залишити заявку")
async def request_start(message: types.Message):
    await Form.fio.set()
    await message.answer("Вкажіть ПІБ:")

@dp.message_handler(state=Form.fio)
async def process_fio(message: types.Message, state: FSMContext):
    await state.update_data(fio=message.text)
    await Form.next()
    await message.answer("Вкажіть адресу підключення:")

@dp.message_handler(state=Form.address)
async def process_address(message: types.Message, state: FSMContext):
    await state.update_data(address=message.text)
    await Form.next()
    await message.answer("Вкажіть номер телефону:")

@dp.message_handler(state=Form.phone)
async def process_phone(message: types.Message, state: FSMContext):
    await state.update_data(phone=message.text)
    await Form.next()
    await message.answer("Опишіть проблему:")

@dp.message_handler(state=Form.problem)
async def process_problem(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    fio = user_data["fio"]
    address = user_data["address"]
    phone = user_data["phone"]
    problem = message.text
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Зберегти в Excel
    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active
    ws.append([now, fio, address, phone, problem])
    wb.save(FILE_NAME)

    await message.answer(
        "✅ Ваша заявка прийнята!\n\n"
        "Якщо проблему не вирішено — телефонуйте 📞 +38 (063) 063-063-6\n"
        "⏰ Пн–Пт з 10:00 до 15:00"
    )
    await state.finish()

# Планувальник: надсилати файл адміну кожні 2 дні
scheduler = AsyncIOScheduler()

def send_excel():
    async def _send():
        if os.path.exists(FILE_NAME):
            today = datetime.now().date()
            start_date = today - timedelta(days=2)
            caption = f"📊 Заявки з {start_date} по {today}"
            await bot.send_document(ADMIN_ID, open(FILE_NAME, "rb"), caption=caption)
    dp.loop.create_task(_send())

scheduler.add_job(send_excel, "interval", days=2)
scheduler.start()

if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
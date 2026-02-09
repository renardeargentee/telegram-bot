import asyncio
import sqlite3
import os
from datetime import datetime

from openpyxl import Workbook
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher(storage=MemoryStorage())

conn = sqlite3.connect("data.db")
cursor = conn.cursor()
cursor.execute("""
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    description TEXT,
    assistant TEXT,
    level INTEGER,
    created_at TEXT
)
""")
conn.commit()


# ---------------- FORMS ----------------
class Form(StatesGroup):
    description = State()
    assistant = State()
    level = State()


class ExportForm(StatesGroup):
    assistant = State()
    date_from = State()
    date_to = State()


class DeleteForm(StatesGroup):
    choose_id = State()


# ---------------- KEYBOARDS ----------------
assistant_keyboard = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="Катерина"), KeyboardButton(text="Авелина")]],
    resize_keyboard=True
)

level_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="1 — мелкий")],
        [KeyboardButton(text="2 — средний")],
        [KeyboardButton(text="3 — серьёзный")]
    ],
    resize_keyboard=True
)

export_assistant_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Все")],
        [KeyboardButton(text="Катерина"), KeyboardButton(text="Авелина")]
    ],
    resize_keyboard=True
)


# ---------------- /start и /help ----------------
@dp.message(Command("start"))
async def start(message: types.Message, state: FSMContext):
    await state.set_state(Form.description)
    await message.answer("Опиши проблему текстом:")


@dp.message(Command("help"))
async def help_command(message: types.Message):
    await message.answer(
        "📖 Справка:\n\n"
        "/start — добавить запись\n"
        "/export — выгрузка в Excel\n"
        "/delete — удалить запись по ID\n"
        "💡 Сначала используй /export, чтобы узнать ID"
    )


# ---------------- ADD RECORD ----------------
@dp.message(Form.description)
async def get_description(message: types.Message, state: FSMContext):
    await state.update_data(description=message.text)
    await state.set_state(Form.assistant)
    await message.answer("Кто?", reply_markup=assistant_keyboard)


@dp.message(Form.assistant)
async def get_assistant(message: types.Message, state: FSMContext):
    if message.text not in ["Катерина", "Авелина"]:
        await message.answer("Выбери кнопкой 👇")
        return
    await state.update_data(assistant=message.text)
    await state.set_state(Form.level)
    await message.answer("Уровень греха?", reply_markup=level_keyboard)


@dp.message(Form.level)
async def get_level(message: types.Message, state: FSMContext):
    levels = {"1 — мелкий": 1, "2 — средний": 2, "3 — серьёзный"}
    if message.text not in levels:
        await message.answer("Выбери кнопкой 👇")
        return
    data = await state.get_data()
    today = datetime.now().strftime("%Y-%m-%d")  # дата без времени
    cursor.execute(
        "INSERT INTO records (description, assistant, level, created_at) VALUES (?, ?, ?, ?)",
        (data["description"], data["assistant"], levels[message.text], today)
    )
    conn.commit()
    await state.clear()
    await message.answer("✅ Принято и сохранено.", reply_markup=ReplyKeyboardRemove())


# ---------------- EXPORT ----------------
@dp.message(Command("export"))
async def export_start(message: types.Message, state: FSMContext):
    await state.set_state(ExportForm.assistant)
    await message.answer("Выбери ассистента:", reply_markup=export_assistant_keyboard)


@dp.message(ExportForm.assistant)
async def export_choose_assistant(message: types.Message, state: FSMContext):
    if message.text not in ["Все", "Катерина", "Авелина"]:
        await message.answer("Выбери кнопкой 👇")
        return
    await state.update_data(assistant=message.text)
    await state.set_state(ExportForm.date_from)
    await message.answer("Введи дату ОТ (в формате ГГГГ-ММ-ДД):", reply_markup=ReplyKeyboardRemove())


@dp.message(ExportForm.date_from)
async def export_date_from(message: types.Message, state: FSMContext):
    await state.update_data(date_from=message.text)
    await state.set_state(ExportForm.date_to)
    await message.answer("Введи дату ДО (в формате ГГГГ-ММ-ДД):")


@dp.message(ExportForm.date_to)
async def export_date_to(message: types.Message, state: FSMContext):
    data = await state.get_data()
    date_from = data["date_from"]
    date_to = message.text
    assistant = data["assistant"]

    if assistant == "Все":
        cursor.execute("""
            SELECT id, description, assistant, level, created_at 
            FROM records 
            WHERE created_at BETWEEN ? AND ?
        """, (date_from, date_to))
    else:
        cursor.execute("""
            SELECT id, description, assistant, level, created_at 
            FROM records 
            WHERE created_at BETWEEN ? AND ?
            AND assistant = ?
        """, (date_from, date_to, assistant))

    rows = cursor.fetchall()

    if not rows:
        await state.clear()
        await message.answer("Нет данных за этот период.")
        return

    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Дата", "Ассистент", "Уровень", "Описание"])  # добавляем колонку ID

    for r in rows:
        ws.append([r[0], r[4], r[2], r[3], r[1]])

    filename = "export.xlsx"
    wb.save(filename)
    await message.answer_document(types.FSInputFile(filename))
    await state.clear()
    import os
    os.remove(filename)


# ---------------- DELETE ----------------
@dp.message(Command("delete"))
async def delete_start(message: types.Message, state: FSMContext):
    await state.set_state(DeleteForm.choose_id)
    await message.answer(
        "Введи ID записи, которую хочешь удалить.\n\n"
        "💡 Сначала используй /export, чтобы узнать ID."
    )


@dp.message(DeleteForm.choose_id)
async def delete_confirm(message: types.Message, state: FSMContext):
    try:
        record_id = int(message.text)
    except ValueError:
        await message.answer("Введи **только число ID**")
        return

    cursor.execute("SELECT * FROM records WHERE id = ?", (record_id,))
    row = cursor.fetchone()
    if not row:
        await message.answer("Запись с таким ID не найдена")
        return

    cursor.execute("DELETE FROM records WHERE id = ?", (record_id,))
    conn.commit()
    await state.clear()
    await message.answer(f"✅ Запись с ID {record_id} удалена")


# ---------------- MAIN ----------------
async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

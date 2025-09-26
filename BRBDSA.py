import logging
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes
import asyncio
from datetime import datetime, date
import os
from openpyxl import Workbook
import json

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# === КОНФИГУРАЦИЯ ===
# ВАЖНО: замените на реальный токен перед запуском
BOT_TOKEN = "8246055725:AAFVR4FK3mRWBb-HVdMSu8S2X8aQMOHoF_Y"

# Система прав доступа через Telegram username
ACCESS_CONFIG = {
    'head': ['ocean_jandal', 'director_username', 'your_username'],
    'manager': ['prlbrlgrl', 'manager2', 'sales1', 'sales2']
}

# Тарифы и подменю (ключи — те же ключи, которые сохраняются в статистике)
TARIFFS = {
    'mts_real': {
        'name': '📱 МТС Риил',
        'submenu': {
            'mts_real_1month': '1 Месяц',
            'mts_real_subscription': 'Абонемент'
        }
    },
    'mts_more': {
        'name': '📶 МТС Больше',
        'submenu': {
            'mts_more_1month': '1 Месяц',
            'mts_more_subscription': 'Абонемент'
        }
    },
    'mts_super': {
        'name': '⚡ МТС Супер',
        'submenu': None
    },
    'membrane': {
        'name': '🛡️ Мембрана',
        'submenu': None
    },
    'yandex_search': {
        'name': '🔍 Яндекс Поиск',
        'submenu': None
    },
    'yandex_x5': {
        'name': '🛒 Яндекс Подписка X5',
        'submenu': {
            'yandex_x5_new': 'Новый клиент',
            'yandex_x5_returning': 'Вернувшийся клиент',
            'yandex_x5_active': 'Действующий клиент'
        }
    },
    'yandex_kids': {
        'name': '👶 Яндекс Подписка Детям',
        'submenu': {
            'yandex_kids_new': 'Новый клиент',
            'yandex_kids_returning': 'Вернувшийся клиент',
            'yandex_kids_active': 'Действующий клиент'
        }
    }
}

# Цены для калькулятора — вы можете менять эти значения позже
PRICE_MAP = {
    'mts_real_1month': 81,
    'mts_real_subscription': 155,
    'mts_more_1month': 81,
    'mts_more_subscription': 175,
    'mts_super': 273,
    'membrane': 494,
    'yandex_search': 180,
    'yandex_x5_new': 650,
    'yandex_x5_returning': 250,
    'yandex_x5_active': 50,
    'yandex_kids_new': 650,
    'yandex_kids_returning': 250,
    'yandex_kids_active': 50
}

class SalesBot:
    def __init__(self, token):
        self.token = token
        self.application = Application.builder().token(token).build()
        self.sales_data = {}  # Храним текущие сессии и продажи в памяти

        # Регистрируем обработчики
        self.application.add_handler(CommandHandler("start", self.start))
        self.application.add_handler(CommandHandler("stats", self.stats))
        self.application.add_handler(CommandHandler("report", self.report))
        self.application.add_handler(CommandHandler("help", self.help_command))
        self.application.add_handler(CommandHandler("id", self.get_id))
        self.application.add_handler(CommandHandler("export", self.export_data))
        self.application.add_handler(CallbackQueryHandler(self.button_handler))

        self.ensure_directories()
        self.migrate_old_data()

    def ensure_directories(self):
        """Создает необходимые директории для хранения данных"""
        os.makedirs('data/daily', exist_ok=True)
        os.makedirs('data/monthly', exist_ok=True)
        os.makedirs('data/backups', exist_ok=True)
    def migrate_old_data(self):
        """Migrate malformed keys in existing JSON stats files.
        Replace keys like 'mts_more_1 Месяц' with 'mts_more_1month' etc based on TARIFFS mapping.
        Creates backups of modified files with .bak extension.
        """
        logger.info("Запуск миграции старых ключей статистики (если есть)...")
        # Build expected keys set
        expected_keys = set(TARIFFS.keys())
        for t_key, t_info in TARIFFS.items():
            submenu = t_info.get('submenu')
            if submenu:
                for sub_key in submenu.keys():
                    expected_keys.add(sub_key)

        # Build mapping from bad human-like keys to correct machine keys
        bad_to_good = {}
        for tariff_key, tariff_info in TARIFFS.items():
            submenu = tariff_info.get('submenu') or {}
            for sub_key, human_name in submenu.items():
                bad_key = f"{tariff_key}_{human_name}"
                bad_to_good[bad_key] = sub_key
                bad_to_good[bad_key.strip()] = sub_key

        def migrate_file(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception as e:
                logger.error(f"Не удалось прочитать {path}: {e}")
                return False
            changed = False
            for user_str, user_data in data.items():
                sales = user_data.get('sales', {}) or {}
                new_sales = {}
                for k, v in sales.items():
                    if k in expected_keys:
                        new_sales[k] = new_sales.get(k, 0) + v
                        continue
                    # direct mapping from bad to good
                    if k in bad_to_good:
                        good = bad_to_good[k]
                        new_sales[good] = new_sales.get(good, 0) + v
                        changed = True
                        continue
                    # heuristic: try to match human-readable suffix to submenu human names
                    if '_' in k:
                        prefix, suffix = k.split('_', 1)
                        if prefix in TARIFFS:
                            submenu = TARIFFS[prefix].get('submenu') or {}
                            found = None
                            for sk, sn in submenu.items():
                                if sn == suffix:
                                    found = sk
                                    break
                            if found:
                                new_sales[found] = new_sales.get(found, 0) + v
                                changed = True
                                continue
                    # fallback - keep original key
                    new_sales[k] = new_sales.get(k, 0) + v
                data[user_str]['sales'] = new_sales
            if changed:
                # backup original file
                try:
                    os.rename(path, path + '.bak')
                except Exception:
                    # if rename fails, continue (no backup)
                    pass
                try:
                    with open(path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    logger.info(f"Мигрирован файл {path}")
                except Exception as e:
                    logger.error(f"Не удалось записать {path}: {e}")
            return changed

        # Apply migration to existing json files
        for folder in ['data/daily', 'data/monthly']:
            if os.path.exists(folder):
                for fname in os.listdir(folder):
                    if fname.lower().endswith('.json'):
                        full = os.path.join(folder, fname)
                        migrate_file(full)
        logger.info("Миграция ключей завершена.")


    def migrate_old_data(self):
        """
        Нормализация старых JSON-файлов статистики, которые могли содержать ключи вида
        'mts_more_1 Месяц' (человеческая часть) вместо 'mts_more_1month' (машинный ключ).
        Эта функция пробегает файлы в data/daily и data/monthly и заменяет такие ключи.
        """
        logger.info("Запуск миграции старых ключей статистики (если есть)...")
        # Собираем ожидаемые ключи и маппинг неправильных -> правильных
        expected_keys = set(TARIFFS.keys())
        for t_key, t_info in TARIFFS.items():
            submenu = t_info.get('submenu')
            if submenu:
                for sub_key in submenu.keys():
                    expected_keys.add(sub_key)

        bad_to_good = {}
        # Заполняем возможные неправильные ключи формата: f"{tariff_key}_{human_name}" -> sub_key
        for tariff_key, tariff_info in TARIFFS.items():
            submenu = tariff_info.get('submenu') or {}
            for sub_key, human_name in submenu.items():
                bad_key = f"{tariff_key}_{human_name}"
                bad_to_good[bad_key] = sub_key
                # also consider trimmed spaces variant
                bad_to_good[bad_key.strip()] = sub_key

        def migrate_file(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception as e:
                logger.error(f"Не удалось прочитать {path}: {e}")
                return False
            changed = False
            for user_str, user_data in data.items():
                sales = user_data.get('sales', {}) or {}
                new_sales = {}
                for k, v in sales.items():
                    if k in expected_keys:
                        new_sales[k] = new_sales.get(k, 0) + v
                        continue
                    if k in bad_to_good:
                        good = bad_to_good[k]
                        new_sales[good] = new_sales.get(good, 0) + v
                        changed = True
                        continue
                    # heuristic: split by first '_' and try to match human name to submenu
                    if '_' in k:
                        prefix, suffix = k.split('_', 1)
                        if prefix in TARIFFS:
                            submenu = TARIFFS[prefix].get('submenu') or {}
                            found = None
                            for sk, sn in submenu.items():
                                if sn == suffix:
                                    found = sk
                                    break
                            if found:
                                new_sales[found] = new_sales.get(found, 0) + v
                                changed = True
                                continue
                    # fallback: keep original key
                    new_sales[k] = new_sales.get(k, 0) + v
                data[user_str]['sales'] = new_sales
            if changed:
                try:
                    with open(path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    logger.info(f"Мигрирован файл {path}")
                except Exception as e:
                    logger.error(f"Не удалось записать {path}: {e}")
            return changed

        for folder in ['data/daily', 'data/monthly']:
            if os.path.exists(folder):
                for fname in os.listdir(folder):
                    if fname.lower().endswith('.json'):
                        full = os.path.join(folder, fname)
                        migrate_file(full)
        logger.info("Миграция ключей завершена.")


    def get_user_role(self, username):
        """Определение роли пользователя по username"""
        if not username:
            return None
        username = username.lower().replace('@', '')
        if username in ACCESS_CONFIG['head']:
            return 'head'
        elif username in ACCESS_CONFIG['manager']:
            return 'manager'
        else:
            return None

    async def start(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда /start"""
        user = update.effective_user
        username = user.username

        role = self.get_user_role(username)
        if not role:
            await update.message.reply_text("❌ Доступ запрещен. Обратитесь к руководителю.")
            return

        if user.id not in self.sales_data:
            self.sales_data[user.id] = {
                'username': username,
                'full_name': user.full_name,
                'role': role,
                'sales': {}
            }

        await self.show_main_menu(update, user.id, role, is_new_message=True)

    async def show_main_menu(self, update, user_id, role, is_new_message=False):
        """Показать главное меню в зависимости от роли"""
        if role == 'manager':
            keyboard = [
                [InlineKeyboardButton("📱 МТС Риил", callback_data="tariff_mts_real")],
                [InlineKeyboardButton("📶 МТС Больше", callback_data="tariff_mts_more")],
                [InlineKeyboardButton("⚡ МТС Супер", callback_data="tariff_mts_super")],
                [InlineKeyboardButton("🛡️ Мембрана", callback_data="tariff_membrane")],
                [InlineKeyboardButton("🔍 Яндекс Поиск", callback_data="tariff_yandex_search")],
                [InlineKeyboardButton("🛒 Яндекс Подписка X5", callback_data="tariff_yandex_x5")],
                [InlineKeyboardButton("👶 Яндекс Подписка Детям", callback_data="tariff_yandex_kids")],
                [
                    InlineKeyboardButton("📊 Статистика за день", callback_data="stats_daily"),
                    InlineKeyboardButton("📈 Общая статистика", callback_data="stats_total")
                ],
                [InlineKeyboardButton("🧮 Калькулятор", callback_data="calculator")]
            ]
            welcome_text = "🎯 Выберите тариф для учета продажи:"
        else:  # head
            keyboard = [
                [InlineKeyboardButton("📊 Статистика за день", callback_data="stats_daily")],
                [InlineKeyboardButton("📈 Общая статистика", callback_data="stats_total")],
                [InlineKeyboardButton("👥 Управление менеджерами", callback_data="manage_managers")],
                [InlineKeyboardButton("🔄 Сброс статистики", callback_data="reset_stats")],
                [InlineKeyboardButton("📤 Экспорт данных", callback_data="export_data")],
                [InlineKeyboardButton("🧮 Калькулятор", callback_data="calculator")]
            ]
            welcome_text = "👑 Панель руководителя:"

        reply_markup = InlineKeyboardMarkup(keyboard)

        # update может быть Update (команда) или CallbackQuery (нажатие кнопки)
        if is_new_message or isinstance(update, Update):
            await update.message.reply_text(welcome_text, reply_markup=reply_markup)
        else:
            # ожидаем, что update — это CallbackQuery
            await update.edit_message_text(welcome_text, reply_markup=reply_markup)

    async def show_tariff_submenu(self, query, tariff_key):
        """Показать подменю для тарифа (если есть)"""
        tariff_info = TARIFFS[tariff_key]

        if tariff_info['submenu']:
            keyboard = []
            # submenu keys уже соответствуют ключам статистики (напр. 'mts_real_1month')
            for sub_key, sub_name in tariff_info['submenu'].items():
                keyboard.append([InlineKeyboardButton(sub_name, callback_data=sub_key)])

            keyboard.append([InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")])

            reply_markup = InlineKeyboardMarkup(keyboard)
            await query.edit_message_text(
                f"🎯 {tariff_info['name']}\nВыберите тип:",
                reply_markup=reply_markup
            )
        else:
            # Тариф без подменю — сразу записываем продажу (ключ равен tariff_key)
            await self.record_sale(query, tariff_key)

    
    async def record_sale(self, query, record_key, display_name: str = None):
        """Запись продажи с сохранением в файлы. Нормализует ключи перед записью."""
        user_id = query.from_user.id
        user_data = self.sales_data[user_id]

        # Normalize potential human-readable keys to machine keys
        record_key_norm = self.normalize_key(record_key)

        # Determine display name
        if not display_name:
            display_name = self.get_display_name(record_key_norm)

        # Update in-memory counters using normalized key
        if record_key_norm not in user_data['sales']:
            user_data['sales'][record_key_norm] = 0
        user_data['sales'][record_key_norm] += 1

        # Save to files using normalized key
        try:
            self.save_daily_sale(user_id, record_key_norm)
            self.save_monthly_sale(user_id, record_key_norm)
        except Exception as e:
            logger.error(f"Ошибка при сохранении: {e}")

        current_count = user_data['sales'][record_key_norm]

        await query.edit_message_text(
            f"✅ Продажа записана!\n\n"
            f"📦 Тариф: {display_name}\n"
            f"📊 Всего продаж этого тарифа: {current_count}\n\n"
            f"Продолжайте в том же духе! 💪"
        )

        await asyncio.sleep(2)
        await self.show_main_menu(query, user_id, user_data['role'])


    def normalize_key(self, key):
        """Попытка нормализации ключа: преобразовать human-like keys в машинные ключи."""
        # direct match
        if not key:
            return key
        if key in TARIFFS:
            return key
        # if matches submenu machine keys
        for t, info in TARIFFS.items():
            submenu = info.get('submenu') or {}
            if key in submenu.keys():
                return key
        # direct match in display_map keys will be resolved by get_display_name, but for backwards compatibility:
        # If key equals a human-readable suffix like '1 Месяц' or 'Абонемент' or construction 'mts_real_1 Месяц',
        # try to map it to corresponding machine key.
        # Case: 'mts_real_1 Месяц' -> split and find matching submenu value
        if '_' in key:
            prefix, suffix = key.split('_', 1)
            if prefix in TARIFFS:
                submenu = TARIFFS[prefix].get('submenu') or {}
                for sub_key, human_name in submenu.items():
                    if human_name == suffix or human_name.strip() == suffix.strip():
                        return sub_key
        # fallback: search any submenu where human_name equals key
        for t, info in TARIFFS.items():
            submenu = info.get('submenu') or {}
            for sub_key, human_name in submenu.items():
                if human_name == key or human_name.strip() == key.strip():
                    return sub_key
        return key


    def save_daily_sale(self, user_id, tariff_key):
        """Сохраняет продажу в дневной файл"""
        today = date.today().strftime('%Y-%m-%d')
        filename = f'data/daily/sales_{today}.json'

        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                daily_data = json.load(f)
        else:
            daily_data = {}

        user_str = str(user_id)
        if user_str not in daily_data:
            daily_data[user_str] = {
                'username': self.sales_data[user_id]['username'],
                'full_name': self.sales_data[user_id]['full_name'],
                'sales': {}
            }

        if tariff_key not in daily_data[user_str]['sales']:
            daily_data[user_str]['sales'][tariff_key] = 0

        daily_data[user_str]['sales'][tariff_key] += 1

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(daily_data, f, ensure_ascii=False, indent=2)

    def save_monthly_sale(self, user_id, tariff_key):
        """Сохраняет продажу в месячный файл"""
        current_month = datetime.now().strftime('%Y-%m')
        filename = f'data/monthly/sales_{current_month}.json'

        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                monthly_data = json.load(f)
        else:
            monthly_data = {}

        user_str = str(user_id)
        if user_str not in monthly_data:
            monthly_data[user_str] = {
                'username': self.sales_data[user_id]['username'],
                'full_name': self.sales_data[user_id]['full_name'],
                'sales': {}
            }

        if tariff_key not in monthly_data[user_str]['sales']:
            monthly_data[user_str]['sales'][tariff_key] = 0

        monthly_data[user_str]['sales'][tariff_key] += 1

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(monthly_data, f, ensure_ascii=False, indent=2)

    def get_daily_stats(self, user_id=None):
        """Получает статистику за текущий день"""
        today = date.today().strftime('%Y-%m-%d')
        filename = f'data/daily/sales_{today}.json'

        if not os.path.exists(filename):
            return {}

        with open(filename, 'r', encoding='utf-8') as f:
            daily_data = json.load(f)

        if user_id:
            user_str = str(user_id)
            return daily_data.get(user_str, {'sales': {}})

        return daily_data

    def get_monthly_stats(self, user_id=None, month=None):
        """Получает статистику за месяц"""
        if not month:
            month = datetime.now().strftime('%Y-%m')

        filename = f'data/monthly/sales_{month}.json'

        if not os.path.exists(filename):
            return {}

        with open(filename, 'r', encoding='utf-8') as f:
            monthly_data = json.load(f)

        if user_id:
            user_str = str(user_id)
            return monthly_data.get(user_str, {'sales': {}})

        return monthly_data

    
    def get_display_name(self, tariff_key):
        """Преобразовать ключ тарифа в читаемое название (устойчиво к человеческим ключам)."""
        display_map = {
            'mts_real_1month': '📱 МТС Риил (1 Месяц)',
            'mts_real_subscription': '📱 МТС Риил (Абонемент)',
            'mts_real': '📱 МТС Риил',
            'mts_more_1month': '📶 МТС Больше (1 Месяц)',
            'mts_more_subscription': '📶 МТС Больше (Абонемент)',
            'mts_more': '📶 МТС Больше',
            'mts_super': '⚡ МТС Супер',
            'membrane': '🛡️ Мембрана',
            'yandex_search': '🔍 Яндекс Поиск',
            'yandex_x5_new': '🛒 Яндекс Подписка X5 (Новый)',
            'yandex_x5_returning': '🛒 Яндекс Подписка X5 (Вернувшийся)',
            'yandex_x5_active': '🛒 Яндекс Подписка X5 (Действующий)',
            'yandex_x5': '🛒 Яндекс Подписка X5',
            'yandex_kids_new': '👶 Яндекс Подписка Детям (Новый)',
            'yandex_kids_returning': '👶 Яндекс Подписка Детям (Вернувшийся)',
            'yandex_kids_active': '👶 Яндекс Подписка Детям (Действующий)',
            'yandex_kids': '👶 Яндекс Подписка Детям'
        }
        # Normalize key first (handles keys like 'mts_more_1 Месяц' etc.)
        try:
            norm = self.normalize_key(tariff_key)
        except Exception:
            norm = tariff_key
        return display_map.get(norm, norm)


    async def show_daily_stats(self, query, user_id):
        """Показать статистику за текущий день для менеджера"""
        user_data = self.sales_data[user_id]
        daily_stats = self.get_daily_stats(user_id)

        if not daily_stats.get('sales'):
            message = f"📊 Статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            message += "📭 Продаж за сегодня еще нет"
        else:
            message = f"📊 Статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            total_sales = 0

            sorted_sales = sorted(daily_stats['sales'].items(), key=lambda x: self.get_display_name(x[0]))
            for tariff_key, count in sorted_sales:
                tariff_name = self.get_display_name(tariff_key)
                message += f"• {tariff_name}: {count} продаж\n"
                total_sales += count

            message += f"\n📈 Всего за день: {total_sales} продаж"

        keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    async def show_total_stats(self, query, user_id):
        """Показать общую статистику за текущий месяц для менеджера"""
        user_data = self.sales_data[user_id]
        monthly_stats = self.get_monthly_stats(user_id)

        if not monthly_stats.get('sales'):
            message = f"📈 Общая статистика за {datetime.now().strftime('%B %Y')}\n\n"
            message += "📭 Продаж за этот месяц еще нет"
        else:
            message = f"📈 Общая статистика за {datetime.now().strftime('%B %Y')}\n\n"
            total_sales = 0

            sorted_sales = sorted(monthly_stats['sales'].items(), key=lambda x: self.get_display_name(x[0]))
            for tariff_key, count in sorted_sales:
                tariff_name = self.get_display_name(tariff_key)
                message += f"• {tariff_name}: {count} продаж\n"
                total_sales += count

            message += f"\n📈 Всего за месяц: {total_sales} продаж"

        keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    async def show_head_daily_stats(self, query):
        """Показать общую дневную статистику для руководителя"""
        daily_stats = self.get_daily_stats()

        if not daily_stats:
            message = f"📊 Общая статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            message += "📭 Продаж за сегодня еще нет"
        else:
            message = f"📊 Общая статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            total_all_sales = 0
            managers_stats = {}

            for user_str, user_data in daily_stats.items():
                manager_name = user_data['full_name']
                managers_stats[manager_name] = {'total': 0, 'details': {}}

                for tariff_key, count in user_data['sales'].items():
                    tariff_name = self.get_display_name(tariff_key)
                    managers_stats[manager_name]['details'][tariff_name] = count
                    managers_stats[manager_name]['total'] += count
                    total_all_sales += count

            for manager, stats in managers_stats.items():
                message += f"👤 {manager}: {stats['total']} продаж\n"
                for tariff, count in stats['details'].items():
                    message += f"   • {tariff}: {count}\n"
                message += "\n"

            message += f"🎯 ОБЩЕЕ КОЛИЧЕСТВО ПРОДАЖ: {total_all_sales}"

        keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    async def show_head_total_stats(self, query):
        """Показать общую месячную статистику для руководителя"""
        monthly_stats = self.get_monthly_stats()

        if not monthly_stats:
            message = f"📈 Общая статистика за {datetime.now().strftime('%B %Y')}\n\n"
            message += "📭 Продаж за этот месяц еще нет"
        else:
            message = f"📈 Общая статистика за {datetime.now().strftime('%B %Y')}\n\n"
            total_all_sales = 0
            managers_stats = {}
            tariff_stats = {}

            for user_str, user_data in monthly_stats.items():
                manager_name = user_data['full_name']
                managers_stats[manager_name] = {'total': 0, 'details': {}}

                for tariff_key, count in user_data['sales'].items():
                    tariff_name = self.get_display_name(tariff_key)

                    # Статистика по тарифам
                    if tariff_name not in tariff_stats:
                        tariff_stats[tariff_name] = 0
                    tariff_stats[tariff_name] += count

                    # Статистика по менеджерам
                    managers_stats[manager_name]['details'][tariff_name] = count
                    managers_stats[manager_name]['total'] += count
                    total_all_sales += count

            message += "👥 СТАТИСТИКА ПО МЕНЕДЖЕРАМ:\n\n"
            for manager, stats in sorted(managers_stats.items()):
                message += f"👤 {manager}: {stats['total']} продаж\n"

            message += "\n📦 СТАТИСТИКА ПО ТАРИФАМ:\n"
            for tariff, count in sorted(tariff_stats.items()):
                message += f"• {tariff}: {count} продаж\n"

            message += f"\n🎯 ОБЩЕЕ КОЛИЧЕСТВО ПРОДАЖ: {total_all_sales}"

        keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    async def show_manage_managers(self, query):
        """Меню управления менеджерами для руководителя"""
        managers = []
        monthly_stats = self.get_monthly_stats()

        for user_str, user_data in monthly_stats.items():
            full_name = user_data.get('full_name', '—')
            total_sales = sum(user_data.get('sales', {}).values())
            managers.append({
                'id': user_str,
                'full_name': full_name,
                'total_sales': total_sales
            })

        keyboard = []
        for manager in managers:
            button_text = f"👤 {manager['full_name']} ({manager['total_sales']} продаж)"
            keyboard.append([InlineKeyboardButton(button_text, callback_data=f"manager_{manager['id']}")])

        keyboard.extend([
            [InlineKeyboardButton("📊 Общая статистика", callback_data="stats_total")],
            [InlineKeyboardButton("🔄 Сбросить всю статистику", callback_data="reset_all")],
            [InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]
        ])

        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text("👥 Управление менеджерами:", reply_markup=reply_markup)

    async def show_reset_options(self, query, manager_id=None):
        """Показать опции сброса статистики"""
        if manager_id:
            manager_data = self.sales_data.get(manager_id)
            if manager_data:
                message = f"🔄 Сброс статистики для {manager_data['full_name']}\n\n"
                keyboard = [
                    [InlineKeyboardButton("🗑️ Сбросить дневную статистику", callback_data=f"reset_daily_{manager_id}")],
                    [InlineKeyboardButton("🗑️ Сбросить месячную статистику", callback_data=f"reset_monthly_{manager_id}")],
                    [InlineKeyboardButton("◀️ Назад к менеджерам", callback_data="manage_managers")]
                ]
            else:
                message = "Пользователь не найден"
                keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        else:
            message = "🔄 Сброс всей статистики\n\n"
            keyboard = [
                [InlineKeyboardButton("🗑️ Сбросить всю дневную статистику", callback_data="reset_all_daily")],
                [InlineKeyboardButton("🗑️ Сбросить всю месячную статистику", callback_data="reset_all_monthly")],
                [InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]
            ]

        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    def reset_daily_stats(self, manager_id=None):
        """Сброс дневной статистики"""
        today = date.today().strftime('%Y-%m-%d')
        filename = f'data/daily/sales_{today}.json'

        if os.path.exists(filename):
            if manager_id:
                with open(filename, 'r', encoding='utf-8') as f:
                    daily_data = json.load(f)

                user_str = str(manager_id)
                if user_str in daily_data:
                    daily_data[user_str]['sales'] = {}

                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(daily_data, f, ensure_ascii=False, indent=2)
            else:
                os.remove(filename)

    def reset_monthly_stats(self, manager_id=None):
        """Сброс месячной статистики"""
        current_month = datetime.now().strftime('%Y-%m')
        filename = f'data/monthly/sales_{current_month}.json'

        if os.path.exists(filename):
            if manager_id:
                with open(filename, 'r', encoding='utf-8') as f:
                    monthly_data = json.load(f)

                user_str = str(manager_id)
                if user_str in monthly_data:
                    monthly_data[user_str]['sales'] = {}

                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(monthly_data, f, ensure_ascii=False, indent=2)
            else:
                os.remove(filename)

    async def export_to_excel(self, query):
        """Экспорт данных в Excel"""
        try:
            wb = Workbook()

            # Дневная статистика
            ws_daily = wb.active
            ws_daily.title = "Дневная статистика"
            headers = ['Менеджер', 'Тариф', 'Количество', 'Дата']
            ws_daily.append(headers)

            today = date.today().strftime('%Y-%m-%d')
            daily_stats = self.get_daily_stats()

            for user_str, user_data in daily_stats.items():
                for tariff_key, count in user_data['sales'].items():
                    tariff_name = self.get_display_name(tariff_key)
                    ws_daily.append([user_data['full_name'], tariff_name, count, today])

            # Месячная статистика
            ws_monthly = wb.create_sheet("Месячная статистика")
            ws_monthly.append(['Менеджер', 'Тариф', 'Количество', 'Месяц'])

            current_month = datetime.now().strftime('%Y-%m')
            monthly_stats = self.get_monthly_stats()

            for user_str, user_data in monthly_stats.items():
                for tariff_key, count in user_data['sales'].items():
                    tariff_name = self.get_display_name(tariff_key)
                    ws_monthly.append([user_data['full_name'], tariff_name, count, current_month])

            filename = f"data/backups/sales_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(filename)

            # Отправляем файл и подтверждение
            await query.message.reply_document(document=open(filename, 'rb'), caption="📊 Экспорт данных продаж")
            await query.edit_message_text("✅ Данные успешно экспортированы в Excel")

        except Exception as e:
            logger.error(f"Ошибка при экспорте в Excel: {e}")
            await query.edit_message_text("❌ Ошибка при экспорте данных")

    async def button_handler(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Обработчик нажатий на кнопки"""
        query = update.callback_query
        await query.answer()

        user_id = query.from_user.id

        if user_id not in self.sales_data:
            await query.edit_message_text("❌ Сессия устарела. Отправьте /start")
            return

        user_data = self.sales_data[user_id]
        data = query.data

        try:
            if data == "back_to_main":
                await self.show_main_menu(query, user_id, user_data['role'])

            elif data == "stats_daily":
                if user_data['role'] == 'head':
                    await self.show_head_daily_stats(query)
                else:
                    await self.show_daily_stats(query, user_id)

            elif data == "stats_total":
                if user_data['role'] == 'head':
                    await self.show_head_total_stats(query)
                else:
                    await self.show_total_stats(query, user_id)

            elif data == "manage_managers":
                await self.show_manage_managers(query)

            elif data == "reset_stats":
                await self.show_reset_options(query)

            elif data.startswith("manager_"):
                manager_id = int(data.split('_')[1])
                await self.show_reset_options(query, manager_id)

            elif data == "reset_all":
                # Удобная точка входа: показать опции полного сброса
                await self.show_reset_options(query)

            elif data.startswith("reset_daily_"):
                if data == "reset_all_daily":
                    self.reset_daily_stats()
                    await query.edit_message_text("✅ Вся дневная статистика сброшена")
                else:
                    manager_id = int(data.split('_')[2])
                    self.reset_daily_stats(manager_id)
                    await query.edit_message_text("✅ Дневная статистика менеджера сброшена")
                await asyncio.sleep(1)
                await self.show_main_menu(query, user_id, user_data['role'])

            elif data.startswith("reset_monthly_"):
                if data == "reset_all_monthly":
                    self.reset_monthly_stats()
                    await query.edit_message_text("✅ Вся месячная статистика сброшена")
                else:
                    manager_id = int(data.split('_')[2])
                    self.reset_monthly_stats(manager_id)
                    await query.edit_message_text("✅ Месячная статистика менеджера сброшена")
                await asyncio.sleep(1)
                await self.show_main_menu(query, user_id, user_data['role'])

            elif data == "export_data":
                await self.export_to_excel(query)

            elif data == "calculator":
                await self.show_calculator(query, user_id)

            # Обработка тарифов — сначала нажатие на кнопку тарифа в главном меню
            elif data.startswith("tariff_"):
                tariff_key = data.replace("tariff_", "")
                if tariff_key in TARIFFS:
                    await self.show_tariff_submenu(query, tariff_key)
                else:
                    # неизвестный тариф
                    await query.edit_message_text("❌ Неизвестный тариф")

            # Если пользователь нажал один из ключей подменю (напр. 'mts_real_1month')
            elif data in self._all_submenu_keys():
                # data здесь — это уже ключ тарифа, который используется как record_key
                await self.record_sale(query, data)

            # Обработка одиночных тарифов, которые не идут через 'tariff_' (на случай)
            elif data in TARIFFS and TARIFFS[data]['submenu'] is None:
                await self.record_sale(query, data)

        except Exception as e:
            logger.error(f"Ошибка в обработчике кнопок: {e}")
            try:
                await query.edit_message_text("❌ Произошла ошибка. Попробуйте снова.")
            except Exception:
                pass

    def _all_submenu_keys(self):
        """Вспомогательная функция — возвращает все ключи подменю"""
        keys = set()
        for t in TARIFFS.values():
            submenu = t.get('submenu')
            if submenu:
                keys.update(submenu.keys())
        return keys

    async def stats(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда /stats - показывает дневную статистику"""
        user = update.effective_user
        user_id = user.id

        if user_id not in self.sales_data:
            await update.message.reply_text("❌ Отправьте /start для начала работы")
            return

        user_data = self.sales_data[user_id]
        daily_stats = self.get_daily_stats(user_id)

        if not daily_stats.get('sales'):
            message = f"📊 Статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            message += "📭 Продаж за сегодня еще нет"
        else:
            message = f"📊 Статистика за сегодня ({date.today().strftime('%d.%m.%Y')})\n\n"
            total_sales = 0

            sorted_sales = sorted(daily_stats['sales'].items(), key=lambda x: self.get_display_name(x[0]))
            for tariff_key, count in sorted_sales:
                tariff_name = self.get_display_name(tariff_key)
                message += f"• {tariff_name}: {count} продаж\n"
                total_sales += count

            message += f"\n📈 Всего за день: {total_sales} продаж"

        await update.message.reply_text(message)

    async def report(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда /report - общий отчет для руководителя"""
        user = update.effective_user
        user_id = user.id

        if user_id not in self.sales_data:
            await update.message.reply_text("❌ Отправьте /start для начала работы")
            return

        user_data = self.sales_data[user_id]

        if user_data['role'] != 'head':
            await update.message.reply_text("❌ Доступ только для руководителя")
            return

        monthly_stats = self.get_monthly_stats()

        if not monthly_stats:
            message = f"📈 Общий отчет за {datetime.now().strftime('%B %Y')}\n\n"
            message += "📭 Продаж за этот месяц еще нет"
        else:
            message = f"📈 Общий отчет за {datetime.now().strftime('%B %Y')}\n\n"
            total_all_sales = 0
            managers_stats = {}

            for user_str, user_data in monthly_stats.items():
                manager_name = user_data['full_name']
                managers_stats[manager_name] = {'total': 0}

                for count in user_data['sales'].values():
                    managers_stats[manager_name]['total'] += count
                    total_all_sales += count

            message += "👥 СТАТИСТИКА ПО МЕНЕДЖЕРАМ:\n\n"
            for manager, stats in sorted(managers_stats.items()):
                message += f"👤 {manager}: {stats['total']} продаж\n"

            message += f"\n🎯 ОБЩЕЕ КОЛИЧЕСТВО ПРОДАЖ: {total_all_sales}"

        await update.message.reply_text(message)

    async def help_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE = None):
        """Команда /help"""
        # показываем справку (вызов либо через команду, либо редактируем сообщение callback'ом)
        help_text = (
            "🤖 *КОМАНДЫ БОТА:*\n\n"
            "/start - Главное меню\n"
            "/stats - Статистика за сегодня\n"
            "/report - Общий отчет (только для руководителя)\n"
            "/export - Экспорт данных в Excel (только для руководителя)\n"
            "/help - Эта справка\n"
            "/id - Показать ваш ID и username\n\n"
            "*ДОСТУПНЫЕ ТАРИФЫ:*\n"
            "• 📱 МТС Риил (1 Месяц/Абонемент)\n"
            "• 📶 МТС Больше (1 Месяц/Абонемент)\n"
            "• ⚡ МТС Супер\n"
            "• 🛡️ Мембрана\n"
            "• 🔍 Яндекс Поиск\n"
            "• 🛒 Яндекс Подписка X5 (Новый/Вернувшийся/Действующий)\n"
            "• 👶 Яндекс Подписка Детям (Новый/Вернувшийся/Действующий)\n\n"
            "*📊 СИСТЕМА СТАТИСТИКИ:*\n"
            "- Дневная статистика (автоматически сохраняется)\n"
            "- Месячная статистика (сохраняется долгосрочно)\n"
            "- Экспорт в Excel для анализа"
        )

        # если это Update (команда) — пришлём as message, иначе редактируем callback'ом
        if isinstance(update, Update):
            await update.message.reply_text(help_text, parse_mode='Markdown')
        else:
            await update.edit_message_text(help_text, parse_mode='Markdown')
            await self.show_main_menu(update, update.from_user.id, self.sales_data[update.from_user.id]['role'])

    async def get_id(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда /id - показать информацию о пользователе"""
        user = update.effective_user
        role = self.get_user_role(user.username)

        message = (
            f"🆔 **Ваша информация:**\n\n"
            f"**ID:** `{user.id}`\n"
            f"**Username:** @{user.username or 'не установлен'}\n"
            f"**Имя:** {user.full_name}\n"
            f"**Роль:** {role or 'не определена'}\n\n"
            "📋 **Для доступа к боту:**\n"
            "1. Сообщите ваш username руководителю\n"
            "2. Руководитель добавит вас в систему\n"
            "3. Перезапустите бот командой /start"
        )

        await update.message.reply_text(message, parse_mode='Markdown')

    async def export_data(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Команда /export для экспорта данных"""
        user = update.effective_user
        user_id = user.id

        if user_id not in self.sales_data:
            await update.message.reply_text("❌ Отправьте /start для начала работы")
            return

        user_data = self.sales_data[user_id]

        if user_data['role'] != 'head':
            await update.message.reply_text("❌ Доступ только для руководителя")
            return

        await update.message.reply_text("📊 Подготовка экспорта данных...")

        # Создаём mock-query для передачи в export_to_excel
        class MockQuery:
            def __init__(self, message):
                self.message = message

        mock_query = MockQuery(update.message)
        await self.export_to_excel(mock_query)

    async def show_calculator(self, query, user_id):
        """Показать расчет суммы всех проданных тарифов, умноженных на PRICE_MAP"""
        user_data = self.sales_data[user_id]

        # Для руководителя — суммируем по всем менеджерам за текущий месяц и показываем детализацию по менеджерам
        if user_data['role'] == 'head':
            monthly_stats = self.get_monthly_stats()
            if not monthly_stats:
                message = "🧮 Калькулятор: данных нет для текущего месяца"
                await query.edit_message_text(message)
                return

            total_all = 0
            message = f"🧮 Калькулятор дохода — {datetime.now().strftime('%B %Y')}\n\n"
            for user_str, udata in monthly_stats.items():
                manager_total = 0
                message += f"👤 {udata.get('full_name','—')}:\n"
                for tariff_key, count in udata.get('sales', {}).items():
                    price = PRICE_MAP.get(tariff_key, 0)
                    subtotal = price * count
                    manager_total += subtotal
                    message += f"   • {self.get_display_name(tariff_key)}: {count} × {price} = {subtotal} ₽\n"
                message += f"   ➜ Всего у менеджера: {manager_total} ₽\n\n"
                total_all += manager_total

            message += f"🎯 ИТОГО ПО ВСЕМ МЕНЕДЖЕРАМ: {total_all} ₽"

        else:
            monthly_stats = self.get_monthly_stats(user_id)
            stats = monthly_stats.get('sales', {}) if monthly_stats else {}

            if not stats:
                message = "🧮 Калькулятор: у вас нет продаж за текущий месяц"
                await query.edit_message_text(message)
                return

            total = 0
            message = f"🧮 Калькулятор дохода — {datetime.now().strftime('%B %Y')}\n\n"
            for tariff_key, count in stats.items():
                price = PRICE_MAP.get(tariff_key, 0)
                subtotal = price * count
                total += subtotal
                message += f"• {self.get_display_name(tariff_key)}: {count} × {price} = {subtotal} ₽\n"

            message += f"\n💰 Итого: {total} ₽"

        keyboard = [[InlineKeyboardButton("◀️ Назад", callback_data="back_to_main")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await query.edit_message_text(message, reply_markup=reply_markup)

    def run(self):
        """Запуск бота"""
        print("=" * 60)
        print("🤖 ЗАПУСК ТЕЛЕГРАМ БОТА С СИСТЕМОЙ СТАТИСТИКИ")
        print("=" * 60)
        print("✅ Токен бота: установлен" if self.token else "❌ Токен не задан")
        print("✅ Система хранения данных: JSON + Excel")
        print("✅ Статистика: Дневная + Месячная")
        print("✅ Функции сброса: Реализованы")
        print("⏳ Ожидание сообщений...")
        print("=" * 60)

        try:
            self.application.run_polling()
        except KeyboardInterrupt:
            print("\n🛑 Бот остановлен")
        except Exception as e:
            print(f"❌ Ошибка: {e}")


# === Запуск бота ===
if __name__ == '__main__':
    # Установите необходимые библиотеки:
    # pip install python-telegram-bot openpyxl

    bot = SalesBot(BOT_TOKEN)
    bot.run()
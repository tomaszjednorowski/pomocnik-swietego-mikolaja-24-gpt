import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import smtplib
import ssl
import re
import threading
import queue
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random

# --- NOWY IMPORT DLA TŁUMACZEŃ ---
try:
    from googletrans import Translator, LANGUAGES
    GOOGLE_TRANS_AVAILABLE = True
except ImportError:
    GOOGLE_TRANS_AVAILABLE = False
    LANGUAGES = {} # Pusta lista, jeśli biblioteka nie jest dostępna


# --- Stałe ---
GMAIL_SMTP_SERVER = "smtp.gmail.com"
GMAIL_SMTP_PORT = 465
OCZEKIWANE_NAGLOWKI = ['imie', 'partner', 'mail', 'jezyk']

# --- ZAKTUALIZOWANY SZABLON (Tylko 2 znaczniki) ---
DEFAULT_SUBJECT = "Tajemniczy Mikołaj - Wyniki Losowania"
DEFAULT_BODY = (
    "Cześć {imie},\n\n"
    "Nadszedł ten magiczny czas! Właśnie odbyło się losowanie Tajemniczego Mikołaja.\n\n"
    "W tym roku Twoim zadaniem jest przygotowanie prezentu dla:\n\n"
    "**{wylosowana_osoba}**\n\n"
    "Zachowaj tę informację w tajemnicy i Wesołych Świąt!\n\n"
    "PS. Pamiętaj o ustalonym budżecie!"
)

# --- Funkcja pomocnicza do walidacji email ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email)

# --- LOGIKA TŁUMACZENIA ---
if GOOGLE_TRANS_AVAILABLE:
    translator = Translator()

def get_translation(text, target_lang, source_lang='pl'):
    if target_lang == source_lang or not text:
        return text
        
    if not GOOGLE_TRANS_AVAILABLE:
        return f"[{target_lang.upper()} - Tłumaczenie dla: {text[:20]}... (Zainstaluj googletrans)]"

    try:
        text_to_translate = text.replace("{imie}", "PLACEHOLDER_NAME").replace("{wylosowana_osoba}", "PLACEHOLDER_DRAWN")
        translated = translator.translate(text_to_translate, src=source_lang, dest=target_lang)
        translated_text = translated.text.replace("PLACEHOLDER_NAME", "{imie}").replace("PLACEHOLDER_DRAWN", "{wylosowana_osoba}")
        return translated_text
    except Exception as e:
        print(f"Błąd tłumaczenia: {e}")
        return f"[{target_lang.upper()}] BŁĄD TŁUMACZENIA: {e}"

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        
        # --- BAZA TŁUMACZEŃ ---
        self.translations = {
            'pl': {
                'lang_name': "Polski (PL)",
                'window_title': "SantaClausMailer",
                'main_title': "SantaClausMailer",
                'footer_author': "Autor: Tomasz Jednorowski & AI | v2.0 (2024-2025)", # ZMIANA
                'tab1_participants': "1. Uczestnicy",
                'tab2_template': "2. Treść Maila i Start",
                'tab3_report': "3. Raport (na żywo)",
                'login_frame_title': "Dane logowania i lista uczestników",
                'email_label': "Twój email Gmail do wysyłki maili:",
                'password_label': "Hasło do aplikacji:",
                'default_email_lang_label': "Domyślny język maila:",
                'test_login_btn': "Przetestuj logowanie",
                'participants_tab_manual': "Wprowadź ręcznie",
                'participants_tab_csv': "Importuj z CSV",
                'manual_name_label': "Imię:",
                'manual_partner_label': "Partner (zostaw puste, jeśli brak):",
                'manual_mail_label': "Mail:",
                'manual_lang_label': "Język maila (domyślny, jeśli pusty):",
                'add_to_list_btn': "Dodaj do listy",
                'participant_list_label': "Lista uczestników",
                'tree_col_name': "Imię",
                'tree_col_partner': "Partner",
                'tree_col_mail': "Mail",
                'tree_col_lang': "Język",
                'edit_btn': "Edytuj zaznaczony",
                'delete_btn': "Usuń zaznaczony",
                'csv_file_label': "Plik CSV:",
                'csv_no_file': "Nie wybrano pliku...",
                'csv_select_btn': "Wybierz plik",
                'csv_instructions_title': "Instrukcja formatu CSV",
                'csv_instr_1': "Plik musi być zapisany w formacie .csv (kodowanie UTF-8).\nDane muszą być oddzielone przecinkami (,).\n\nPlik musi zawierać DOKŁADNIE 4 nagłówki w pierwszym wierszu:\nimie,partner,mail,jezyk\n\n•  Nagłówki muszą być w tej kolejności i bez polskich znaków.\n•  imię: Unikalne imię uczestnika (wymagane).\n•  partner: Imię partnera do wykluczenia (zostaw puste, jeśli brak).\n•  mail: Adres e-mail uczestnika (wymagany).\n•  jezyk: Kod języka (np. pl, sk, en, de). Użyje domyślnego języka, jeśli puste.",
                'template_frame_title': "2.1 Treść Maila (Szablon)",
                'template_subject_label': "Temat:",
                'template_body_label': "Treść (Język źródłowy: Polski):",
                'tags_label': "Dostępne znaczniki:",
                'tag_btn_imie': "Twoje imię",
                'tag_btn_wylosowana': "Wylosowana Osoba",
                'reset_btn': "Przywróć domyślną treść",
                'start_frame_title': "2.2 Uruchomienie",
                'show_results_check': "Pokaż mi wyniki losowania w logach (dla administratora)",
                'start_btn': "Rozpocznij Losowanie i Wysyłkę",
                'report_frame_title': "Raport z Losowania (na żywo)",
                'clear_log_btn': "Wyczyść logi",
                'preview_tab': "Podgląd i Edycja Tłumaczeń",
                'generate_preview_btn': "Generuj/Odśwież Podgląd i Edycję",
                'default_subject': "Tajemniczy Mikołaj - Wyniki Losowania",
                'default_body': "Cześć {imie},\n\nNadszedł ten magiczny czas! Właśnie odbyło się losowanie Tajemniczego Mikołaja.\n\nW tym roku Twoim zadaniem jest przygotowanie prezentu dla:\n\n**{wylosowana_osoba}**\n\nZachowaj tę informację w tajemnicy i Wesołych Świąt!\n\nPS. Pamiętaj o ustalonym budżecie!",
                'confirm_title': "Potwierdzenie",
                'confirm_reset_template': "Czy na pewno chcesz zresetować treść maila do domyślnej?",
                'log_template_reset': "[INFO] Zresetowano szablon maila do domyślnego.",
                'warn_title': "Puste pola",
                'warn_empty_fields': "Pola 'Imię' i 'Mail' są wymagane.",
                'warn_validation_title': "Błąd walidacji",
                'warn_invalid_email': "Wprowadzony adres e-mail jest nieprawidłowy.",
                'log_edit_loaded': "Załadowano '{0}' do edycji. Popraw dane i dodaj rekord ponownie do listy.",
                'warn_no_selection_title': "Brak zaznaczenia",
                'warn_no_selection_edit': "Zaznacz rekord, który chcesz edytować.",
                'warn_too_many_selected': "Zaznacz tylko jeden rekord na raz, aby edytować.",
                'warn_no_selection_delete': "Zaznacz rekord (lub rekordy), które chcesz usunąć.",
                'confirm_delete_records': "Czy na pewno chcesz usunąć {0} rekord(ów)?",
                'log_file_selected': "Wybrano plik: {0}",
                'log_insert_tag_error': "[BŁĄD] Nie można wstawić znacznika. Błąd pola tekstowego.",
                'confirm_clear_logs': "Czy na pewno chcesz wyczyścić okno raportu?",
                'log_logs_cleared': "[INFO] Logi zostały wyczyszczone.",
                'error_title': "Błąd wejścia",
                'error_login_fields_empty': "Wypełnij pola: Email i Hasło administratora, aby przetestować.",
                'log_login_test_start': "[INFO] Rozpoczynam test logowania...",
                'log_login_success': "[SUKCES] Logowanie pomyślne! Email i hasło są poprawne.",
                'log_login_fail_auth': "[BŁĄD] Logowanie nieudane. Sprawdź email lub hasło do aplikacji.",
                'log_login_fail_connection': "[BŁĄD] Wystąpił błąd połączenia lub inny: {0}",
                'error_login_fields_main': "Wypełnij pola: Email i Hasło administratora.",
                'error_template_empty': "Temat i treść maila nie mogą być puste.",
                'error_no_file_selected': "Wybierz plik CSV.",
                'error_file_empty': "Plik CSV jest pusty.",
                'error_csv_header': "Nieprawidłowe nagłówki w CSV.\nOczekiwano: {0}\nZnaleziono: {1}",
                'error_csv_read': "Nie można odczytać pliku CSV: {0}",
                'error_participant_list_empty': "Lista uczestników jest pusta. Dodaj kogoś.",
                'error_participant_list_short': "Musi być przynajmniej 2 uczestników, aby przeprowadzić losowanie.",
                'log_draw_start': "--- Rozpoczynanie procesu losowania ---",
                'log_btn_starting': "Losowanie i Wysyłanie...",
                'log_btn_idle': "Rozpocznij Losowanie i Wysyłkę",
                'log_validating_participants': "Walidowanie {0} uczestników...",
                'log_err_columns': "[BŁĄD Rekord {0}] Nieprawidłowa liczba kolumn. Pomijam.",
                'log_err_empty_row': "[BŁĄD Rekord {0}] Puste 'Imię' ('{1}') lub 'Mail' ('{2}'). Pomijam.",
                'log_err_invalid_email': "[BŁĄD Rekord {0}] Nieprawidłowy email: '{1}'. Pomijam.",
                'log_err_duplicate_name': "[BŁĄD Rekord {0}] Zduplikowane imię po normalizacji: '{1}' (z '{2}'). Imię musi być unikalne! Przerywam.",
                'log_err_validation_short': "[BŁĄD] Po walidacji zostało mniej niż 2 uczestników. Przerywam.",
                'log_validation_success': "[INFO] Pomyślnie zwalidowano {0} uczestników.",
                'log_draw_starting': "[INFO] Rozpoczynam losowanie...",
                'log_err_draw_failed': "[BŁĄD KRYTYCZNY] Nie udało się znaleźć poprawnego losowania po 1000 prób.\n[INFO] Może to być niemożliwe (np. 3 osoby w parach). Spróbuj ponownie lub zmień listę.",
                'log_draw_success': "[SUKCES] Losowanie zakończone pomyślnie! Przygotowuję wysyłkę...",
                'log_show_results_title': "--- WYNIKI LOSOWANIA (Tryb podglądu) ---",
                'log_show_results_pair': "  {0} -> {1}",
                'log_show_results_footer': "-----------------------------------------",
                'log_connecting_gmail': "Łączenie z serwerem Gmail...",
                'log_login_success_main': "[SUKCES] Zalogowano pomyślnie jako {0}",
                'log_login_fail_main': "[BŁĄD KRYTYCZNY] Nieprawidłowy email lub hasło aplikacji.",
                'log_login_fail_connection_main': "[BŁĄD KRYTYCZNY] Błąd logowania: {0}",
                'log_processing_send': "[Procesowanie] Wysyłanie do {0} ({1}) w języku {2}...",
                'log_err_template_key': "  [BŁĄD] Błąd w szablonie maila. Brak klucza: {0}. Użyj {{imie}} lub {{wylosowana_osoba}}.",
                'log_send_success': "  [SUKCES] Wysłano informację do {0}",
                'log_send_fail': "  [BŁĄD] Nie wysłano do {0}: {1}",
                'log_summary_title': "--- Podsumowanie Wysyłki ---",
                'log_summary_success': "Informacje wysłane pomyślnie: {0}",
                'log_summary_errors': "Błędy: {0}",
                'log_err_critical': "[BŁĄD KRYTYCZNY] Wystąpił nieoczekiwany błąd: {0}",
                'log_process_finished': "--- ZAKOŃCZONO PROCES ---",
                'log_generating_previews': "[INFO] Generowanie podglądu tłumaczeń...",
                'log_previews_done': "[INFO] Podgląd gotowy.",
                'log_err_googletrans_missing': "[BŁĄD] Biblioteka 'googletrans' nie jest zainstalowana.\nUruchom 'pip install googletrans==4.0.0-rc1' w terminalu,\naby włączyć tłumaczenie."
            },
            'sk': {
                'lang_name': "Slovenčina (SK)",
                'footer_author': "Autor: Tomasz Jednorowski & AI | v2.0 (2024-2025)", # ZMIANA
                # ... (reszta tłumaczeń sk bez zmian) ...
                'window_title': "SantaClausMailer",
                'main_title': "SantaClausMailer",
                'tab1_participants': "1. Účastníci",
                'tab2_template': "2. Obsah Mailu a Štart",
                'tab3_report': "3. Report (naživo)",
                'login_frame_title': "Prihlasovacie údaje a zoznam účastníkov",
                'email_label': "Váš Gmail na odoslanie e-mailov:",
                'password_label': "Heslo aplikácie:",
                'default_email_lang_label': "Predvolený jazyk e-mailu:",
                'test_login_btn': "Otestovať prihlásenie",
                'participants_tab_manual': "Zadať ručne",
                'participants_tab_csv': "Importovať z CSV",
                'manual_name_label': "Meno:",
                'manual_partner_label': "Partner (nechajte prázdne, ak žiadny):",
                'manual_mail_label': "Mail:",
                'manual_lang_label': "Jazyk mailu (predvolený, ak je prázdny):",
                'add_to_list_btn': "Pridať do zoznamu",
                'participant_list_label': "Zoznam účastníkov",
                'tree_col_name': "Meno",
                'tree_col_partner': "Partner",
                'tree_col_mail': "Mail",
                'tree_col_lang': "Jazyk",
                'edit_btn': "Upraviť označené",
                'delete_btn': "Odstrániť označené",
                'csv_file_label': "Súbor CSV:",
                'csv_no_file': "Nevybraný žiadny súbor...",
                'csv_select_btn': "Vybrať súbor",
                'csv_instructions_title': "Inštrukcie pre formát CSV",
                'csv_instr_1': "Súbor musí byť uložený vo formáte .csv (kódovanie UTF-8).\nÚdaje musia byť oddelené čiarkami (,).\n\nSúbor musí obsahovať PRESNE 4 hlavičky v prvom riadku:\nimie,partner,mail,jezyk\n\n•  Hlavičky musia byť v tomto poradí a bez diakritiky.\n•  imie: Unikátne meno účastníka (povinné).\n•  partner: Meno partnera na vylúčenie (nechajte prázdne, ak žiadny).\n•  mail: E-mailová adresa účastníka (povinné).\n•  jezyk: Kód jazyka (napr. pl, sk, en, de). Ak je prázdne, použije sa predvolený jazyk.",
                'template_frame_title': "2.1 Obsah Mailu (Šablóna)",
                'template_subject_label': "Predmet:",
                'template_body_label': "Obsah (Zdrojový jazyk: Poľština):",
                'tags_label': "Dostupné značky:",
                'tag_btn_imie': "Tvoje meno",
                'tag_btn_wylosowana': "Vylosovaná Osoba",
                'reset_btn': "Obnoviť predvolený obsah",
                'start_frame_title': "2.2 Spustenie",
                'show_results_check': "Ukáž mi výsledky žrebovania v logoch (pre administrátora)",
                'start_btn': "Spustiť Žrebovanie a Odoslanie",
                'report_frame_title': "Report zo Žrebovania (naživo)",
                'clear_log_btn': "Vyčistiť logy",
                'preview_tab': "Náhľad a Úprava Prekladov",
                'generate_preview_btn': "Generovať/Obnoviť Náhľad Prekladov",
                'default_subject': "Tajný Mikuláš - Výsledky Žrebovania",
                'default_body': "Ahoj {imie},\n\nNastal ten čarovný čas! Práve prebehlo žrebovanie Tajného Mikuláša.\n\nTvojou úlohou tento rok je pripraviť darček pre:\n\n**{wylosowana_osoba}**\n\nTúto informáciu si nechaj v tajnosti a prajeme veselé Vianoce!\n\nPS. Nezabudni na dohodnutý rozpočet!",
                'confirm_title': "Potvrdenie",
                'confirm_reset_template': "Naozaj chcete obnoviť obsah e-mailu na predvolený?",
                'log_template_reset': "[INFO] Šablóna e-mailu bola obnovená na predvolenú.",
                'warn_title': "Prázdne polia",
                'warn_empty_fields': "Polia 'Meno' a 'Mail' sú povinné.",
                'warn_validation_title': "Chyba validácie",
                'warn_invalid_email': "Zadaná e-mailová adresa je neplatná.",
                'log_edit_loaded': "Nahrávam '{0}' na úpravu. Opravte údaje a pridajte záznam znova do zoznamu.",
                'warn_no_selection_title': "Nič nie je označené",
                'warn_no_selection_edit': "Označte záznam, ktorý chcete upraviť.",
                'warn_too_many_selected': "Označte iba jeden záznam naraz, aby ste ho mohli upraviť.",
                'warn_no_selection_delete': "Označte záznam (alebo záznamy), ktoré chcete odstrániť.",
                'confirm_delete_records': "Naozaj chcete odstrániť {0} záznam(ov)?",
                'log_file_selected': "Vybraný súbor: {0}",
                'log_insert_tag_error': "[CHYBA] Nie je možné vložil značku. Chyba textového poľa.",
                'confirm_clear_logs': "Naozaj chcete vyčistiť okno reportu?",
                'log_logs_cleared': "[INFO] Logy boli vyčistené.",
                'error_title': "Chyba vstupu",
                'error_login_fields_empty': "Vyplňte polia: Email a Heslo administrátora, aby ste mohli testovať.",
                'log_login_test_start': "[INFO] Spúšťam test prihlásenia...",
                'log_login_success': "[ÚSPECH] Prihlásenie úspešné! Email a heslo sú správne.",
                'log_login_fail_auth': "[CHYBA] Prihlásenie zlyhalo. Skontrolujte email alebo heslo aplikácie.",
                'log_login_fail_connection': "[CHYBA] Vyskytla sa chyba pripojenia alebo iná chyba: {0}",
                'error_login_fields_main': "Vyplňte polia: Email a Heslo administrátora.",
                'error_template_empty': "Predmet a obsah e-mailu nesmú byť prázdne.",
                'error_no_file_selected': "Vyberte súbor CSV.",
                'error_file_empty': "Súbor CSV je prázdny.",
                'error_csv_header': "Neplatné hlavičky v CSV.\nOčakávané: {0}\nNájdené: {1}",
                'error_csv_read': "Nie je možné prečítať súbor CSV: {0}",
                'error_participant_list_empty': "Zoznam účastníkov je prázdny. Pridajte niekoho.",
                'error_participant_list_short': "Musia byť aspoň 2 účastníci, aby sa mohlo uskutočniť žrebovanie.",
                'log_draw_start': "--- Spúšťanie procesu žrebovania ---",
                'log_btn_starting': "Žrebovanie a Odosielanie...",
                'log_btn_idle': "Spustiť Žrebovanie a Odoslanie",
                'log_validating_participants': "Validujem {0} účastníkov...",
                'log_err_columns': "[CHYBA Záznam {0}] Nesprávny počet stĺpcov. Preskakujem.",
                'log_err_empty_row': "[CHYBA Záznam {0}] Prázdne 'Meno' ('{1}') alebo 'Mail' ('{2}'). Preskakujem.",
                'log_err_invalid_email': "[CHYBA Záznam {0}] Neplatný email: '{1}'. Preskakujem.",
                'log_err_duplicate_name': "[CHYBA Záznam {0}] Duplicitné meno po normalizácii: '{1}' (z '{2}'). Meno musí byť unikátne! Prerušujem.",
                'log_err_validation_short': "[CHYBA] Po validácii zostali menej ako 2 účastníci. Prerušujem.",
                'log_validation_success': "[INFO] Úspešne validovaných {0} účastníkov.",
                'log_draw_starting': "[INFO] Spúšťam žrebovanie...",
                'log_err_draw_failed': "[KRITICKÁ CHYBA] Nepodarilo sa nájsť platné žrebovanie po 1000 pokusoch.\n[INFO] Je možné, že je to nemožné (napr. 3 ľudia v pároch). Skúste znova alebo zmeňte zoznam.",
                'log_draw_success': "[ÚSPECH] Žrebovanie úspešne ukončené! Pripravujem odosielanie...",
                'log_show_results_title': "--- VÝSLEDKY ŽREBOVANIA (Režim náhľadu) ---",
                'log_show_results_pair': "  {0} -> {1}",
                'log_show_results_footer': "-----------------------------------------",
                'log_connecting_gmail': "Pripájam sa k serveru Gmail...",
                'log_login_success_main': "[ÚSPECH] Úspešne prihlásený ako {0}",
                'log_login_fail_main': "[KRITICKÁ CHYBA] Neplatný email alebo heslo aplikácie.",
                'log_login_fail_connection_main': "[KRITICKÁ CHYBA] Chyba prihlásenia: {0}",
                'log_processing_send': "[Spracúvam] Odosielam {0} ({1}) v jazyku {2}...",
                'log_err_template_key': "  [CHYBA] Chyba v šablóne e-mailu. Chýbajúci kľúč: {0}. Použite {{imie}} alebo {{wylosowana_osoba}}.",
                'log_send_success': "  [ÚSPECH] Informácia odoslaná {0}",
                'log_send_fail': "  [CHYBA] Nepodarilo sa odoslať {0}: {1}",
                'log_summary_title': "--- Súhrn Odosielania ---",
                'log_summary_success': "Informácie úspešne odoslané: {0}",
                'log_summary_errors': "Chyby: {0}",
                'log_err_critical': "[KRITICKÁ CHYBA] Vyskytla sa neočakávaná chyba: {0}",
                'log_process_finished': "--- PROCES UKONČENÝ ---",
                'log_generating_previews': "[INFO] Generujem náhľad prekladov...",
                'log_previews_done': "[INFO] Náhľad hotový.",
                'log_err_googletrans_missing': "[CHYBA] Knižnica 'googletrans' nie je nainštalovaná.\nSpustite 'pip install googletrans==4.0.0-rc1' v termináli,\naby ste povolili preklad."
            },
            'en': {
                'lang_name': "English (EN)",
                'footer_author': "Author: Tomasz Jednorowski & AI | v2.0 (2024-2025)", # ZMIANA
                # ... (reszta tłumaczeń en bez zmian) ...
                'window_title': "SantaClausMailer",
                'main_title': "SantaClausMailer",
                'tab1_participants': "1. Participants",
                'tab2_template': "2. Email Content & Start",
                'tab3_report': "3. Report (Live)",
                'login_frame_title': "Login Details & Participant List",
                'email_label': "Your Gmail for sending emails:",
                'password_label': "Application Password:",
                'default_email_lang_label': "Default Email Language:",
                'test_login_btn': "Test Login",
                'participants_tab_manual': "Enter Manually",
                'participants_tab_csv': "Import from CSV",
                'manual_name_label': "Name:",
                'manual_partner_label': "Partner (leave blank if none):",
                'manual_mail_label': "Email:",
                'manual_lang_label': "Email Language (default if empty):",
                'add_to_list_btn': "Add to List",
                'participant_list_label': "Participant List",
                'tree_col_name': "Name",
                'tree_col_partner': "Partner",
                'tree_col_mail': "Email",
                'tree_col_lang': "Language",
                'edit_btn': "Edit Selected",
                'delete_btn': "Delete Selected",
                'csv_file_label': "CSV File:",
                'csv_no_file': "No file selected...",
                'csv_select_btn': "Select File",
                'csv_instructions_title': "CSV Format Instructions",
                'csv_instr_1': "File must be saved in .csv format (UTF-8 encoding).\nData must be separated by commas (,).\n\nFile must contain EXACTLY 4 headers in the first row:\nimie,partner,mail,jezyk\n\n• Headers must be in this order and without diacritics.\n• imie: Unique participant name (required).\n• partner: Partner's name to exclude (leave blank if none).\n• mail: Participant's email address (required).\n• jezyk: Language code (e.g., pl, sk, en, de). Uses default language if empty.",
                'template_frame_title': "2.1 Email Content (Template)",
                'template_subject_label': "Subject:",
                'template_body_label': "Body (Source Language: Polish):",
                'tags_label': "Available tags:",
                'tag_btn_imie': "Your Name",
                'tag_btn_wylosowana': "Drawn Person",
                'reset_btn': "Reset to Default",
                'start_frame_title': "2.2 Run",
                'show_results_check': "Show me the draw results in the logs (for admin)",
                'start_btn': "Start Draw & Send Emails",
                'report_frame_title': "Draw Report (Live)",
                'clear_log_btn': "Clear Logs",
                'preview_tab': "Translation Preview & Edit",
                'generate_preview_btn': "Generate/Refresh Translation Preview",
                'default_subject': "Secret Santa - Draw Results",
                'default_body': "Hi {imie},\n\nIt's that magical time! The Secret Santa draw has just taken place.\n\nThis year, your task is to prepare a gift for:\n\n**{wylosowana_osoba}**\n\nKeep this information secret and Merry Christmas!\n\nPS. Remember the agreed-upon budget!",
                'confirm_title': "Confirmation",
                'confirm_reset_template': "Are you sure you want to reset the email content to default?",
                'log_template_reset': "[INFO] Email template has been reset to default.",
                'warn_title': "Empty Fields",
                'warn_empty_fields': "'Name' and 'Email' fields are required.",
                'warn_validation_title': "Validation Error",
                'warn_invalid_email': "The entered e-mail address is invalid.",
                'log_edit_loaded': "Loaded '{0}' for editing. Correct the data and add the record to the list again.",
                'warn_no_selection_title': "No Selection",
                'warn_no_selection_edit': "Select a record you want to edit.",
                'warn_too_many_selected': "Select only one record at a time to edit.",
                'warn_no_selection_delete': "Select the record(s) you want to delete.",
                'confirm_delete_records': "Are you sure you want to delete {0} record(s)?",
                'log_file_selected': "Selected file: {0}",
                'log_insert_tag_error': "[ERROR] Cannot insert tag. Text field error.",
                'confirm_clear_logs': "Are you sure you want to clear the report window?",
                'log_logs_cleared': "[INFO] Logs have been cleared.",
                'error_title': "Input Error",
                'error_login_fields_empty': "Fill in the Email and Admin Password fields to test.",
                'log_login_test_start': "[INFO] Starting login test...",
                'log_login_success': "[SUCCESS] Login successful! Email and password are correct.",
                'log_login_fail_auth': "[ERROR] Login failed. Check email or application password.",
                'log_login_fail_connection': "[ERROR] A connection error or other error occurred: {0}",
                'error_login_fields_main': "Fill in the Email and Admin Password fields.",
                'error_template_empty': "Email subject and body cannot be empty.",
                'error_no_file_selected': "Select a CSV file.",
                'error_file_empty': "The CSV file is empty.",
                'error_csv_header': "Invalid headers in CSV.\nExpected: {0}\nFound: {1}",
                'error_csv_read': "Could not read CSV file: {0}",
                'error_participant_list_empty': "Participant list is empty. Add someone.",
                'error_participant_list_short': "There must be at least 2 participants to conduct the draw.",
                'log_draw_start': "--- Starting draw process ---",
                'log_btn_starting': "Drawing & Sending...",
                'log_btn_idle': "Start Draw & Send Emails",
                'log_validating_participants': "Validating {0} participants...",
                'log_err_columns': "[ERROR Record {0}] Invalid column count. Skipping.",
                'log_err_empty_row': "[ERROR Record {0}] Empty 'Name' ('{1}') or 'Email' ('{2}'). Skipping.",
                'log_err_invalid_email': "[ERROR Record {0}] Invalid email: '{1}'. Skipping.",
                'log_err_duplicate_name': "[ERROR Record {0}] Duplicate name after normalization: '{1}' (from '{2}'). Name must be unique! Aborting.",
                'log_err_validation_short': "[ERROR] Fewer than 2 participants left after validation. Aborting.",
                'log_validation_success': "[INFO] Successfully validated {0} participants.",
                'log_draw_starting': "[INFO] Starting draw...",
                'log_err_draw_failed': "[CRITICAL ERROR] Failed to find a valid draw after 1000 tries.\n[INFO] This might be impossible (e.g., 3 people in pairs). Try again or change the list.",
                'log_draw_success': "[SUCCESS] Draw completed successfully! Preparing to send emails...",
                'log_show_results_title': "--- DRAW RESULTS (Admin preview) ---",
                'log_show_results_pair': "  {0} -> {1}",
                'log_show_results_footer': "-----------------------------------------",
                'log_connecting_gmail': "Connecting to Gmail server...",
                'log_login_success_main': "[SUCCESS] Logged in successfully as {0}",
                'log_login_fail_main': "[CRITICAL ERROR] Invalid email or application password.",
                'log_login_fail_connection_main': "[CRITICAL ERROR] Login error: {0}",
                'log_processing_send': "[Processing] Sending to {0} ({1}) in {2}...",
                'log_err_template_key': "  [ERROR] Error in email template. Missing key: {0}. Use {{imie}} or {{wylosowana_osoba}}.",
                'log_send_success': "  [SUCCESS] Sent information to {0}",
                'log_send_fail': "  [ERROR] Failed to send to {0}: {1}",
                'log_summary_title': "--- Send Summary ---",
                'log_summary_success': "Information sent successfully: {0}",
                'log_summary_errors': "Errors: {0}",
                'log_err_critical': "[CRITICAL ERROR] An unexpected error occurred: {0}",
                'log_process_finished': "--- PROCESS FINISHED ---",
                'log_generating_previews': "[INFO] Generating translation previews...",
                'log_previews_done': "[INFO] Previews ready.",
                'log_err_googletrans_missing': "[ERROR] 'googletrans' library not installed.\nRun 'pip install googletrans==4.0.0-rc1' in your terminal\nto enable translation."
            },
            'de': {
                'lang_name': "Deutsch (DE)",
                'footer_author': "Autor: Tomasz Jednorowski & AI | v2.0 (2024-2025)", # ZMIANA
                # ... (reszta tłumaczeń de bez zmian) ...
                'window_title': "SantaClausMailer",
                'main_title': "SantaClausMailer",
                'tab1_participants': "1. Teilnehmer",
                'tab2_template': "2. E-Mail-Inhalt & Start",
                'tab3_report': "3. Bericht (Live)",
                'login_frame_title': "Anmeldedaten & Teilnehmerliste",
                'email_label': "Ihre Gmail zum Senden von E-Mails:",
                'password_label': "Anwendungspasswort:",
                'default_email_lang_label': "Standard-E-Mail-Sprache:",
                'test_login_btn': "Anmeldung testen",
                'participants_tab_manual': "Manuell eingeben",
                'participants_tab_csv': "Aus CSV importieren",
                'manual_name_label': "Name:",
                'manual_partner_label': "Partner (leer lassen, falls keiner):",
                'manual_mail_label': "E-Mail:",
                'manual_lang_label': "E-Mail-Sprache (Standard, falls leer):",
                'add_to_list_btn': "Zur Liste hinzufügen",
                'participant_list_label': "Teilnehmerliste",
                'tree_col_name': "Name",
                'tree_col_partner': "Partner",
                'tree_col_mail': "E-Mail",
                'tree_col_lang': "Sprache",
                'edit_btn': "Auswahl bearbeiten",
                'delete_btn': "Auswahl löschen",
                'csv_file_label': "CSV-Datei:",
                'csv_no_file': "Keine Datei ausgewählt...",
                'csv_select_btn': "Datei auswählen",
                'csv_instructions_title': "Anleitung CSV-Format",
                'csv_instr_1': "Datei muss im .csv-Format (UTF-8) gespeichert sein.\nDaten müssen durch Kommas (,) getrennt sein.\n\nDie Datei muss GENAU 4 Kopfzeilen in der ersten Zeile enthalten:\nimie,partner,mail,jezyk\n\n• Die Kopfzeilen müssen in dieser Reihenfolge und ohne Umlaute sein.\n• imie: Eindeutiger Name des Teilnehmers (erforderlich).\n• partner: Name des Partners, der ausgeschlossen werden soll (leer lassen, falls keiner).\n• mail: E-Mail-Adresse des Teilnehmers (erforderlich).\n• jezyk: Sprachcode (z.B. pl, sk, en, de). Verwendet Standardsprache, falls leer.",
                'template_frame_title': "2.1 E-Mail-Inhalt (Vorlage)",
                'template_subject_label': "Betreff:",
                'template_body_label': "Inhalt (Quellsprache: Polnisch):",
                'tags_label': "Verfügbare Platzhalter:",
                'tag_btn_imie': "Dein Name",
                'tag_btn_wylosowana': "Gezogene Person",
                'reset_btn': "Auf Standard zurücksetzen",
                'start_frame_title': "2.2 Starten",
                'show_results_check': "Auslosungsergebnisse im Log anzeigen (für Admin)",
                'start_btn': "Auslosung starten & E-Mails senden",
                'report_frame_title': "Auslosungsbericht (Live)",
                'clear_log_btn': "Logs löschen",
                'preview_tab': "Vorschau & Bearbeitung",
                'generate_preview_btn': "Vorschau generieren/aktualisieren",
                'default_subject': "Wichteln - Auslosungsergebnisse",
                'default_body': "Hallo {imie},\n\nes ist wieder so weit! Die Wichtel-Auslosung hat gerade stattgefunden.\n\nDieses Jahr ist es deine Aufgabe, ein Geschenk für:\n\n**{wylosowana_osoba}**\n\nzu besorgen.\n\nBehalte diese Information für dich und frohe Weihnachten!\n\nPS. Denk an das vereinbarte Budget!",
                'confirm_title': "Bestätigung",
                'confirm_reset_template': "Möchten Sie den E-Mail-Inhalt wirklich auf den Standard zurücksetzen?",
                'log_template_reset': "[INFO] E-Mail-Vorlage wurde auf Standard zurückgesetzt.",
                'warn_title': "Leere Felder",
                'warn_empty_fields': "Die Felder 'Name' und 'E-Mail' sind erforderlich.",
                'warn_validation_title': "Validierungsfehler",
                'warn_invalid_email': "Die eingegebene E-Mail-Adresse ist ungültig.",
                'log_edit_loaded': "'{0}' zum Bearbeiten geladen. Korrigieren Sie die Daten und fügen Sie den Datensatz erneut zur Liste hinzu.",
                'warn_no_selection_title': "Keine Auswahl",
                'warn_no_selection_edit': "Wählen Sie einen Datensatz aus, den Sie bearbeiten möchten.",
                'warn_too_many_selected': "Wählen Sie nur einen Datensatz auf einmal aus, um ihn zu bearbeiten.",
                'warn_no_selection_delete': "Wählen Sie den/die Datensatz/Datensätze aus, den/die Sie löschen möchten.",
                'confirm_delete_records': "Möchten Sie wirklich {0} Datensatz/Datensätze löschen?",
                'log_file_selected': "Datei ausgewählt: {0}",
                'log_insert_tag_error': "[FEHLER] Platzhalter konnte nicht eingefügt werden. Textfeld-Fehler.",
                'confirm_clear_logs': "Möchten Sie das Berichtsfenster wirklich leeren?",
                'log_logs_cleared': "[INFO] Logs wurden gelöscht.",
                'error_title': "Eingabefehler",
                'error_login_fields_empty': "Füllen Sie die Felder E-Mail und Admin-Passwort aus, um zu testen.",
                'log_login_test_start': "[INFO] Starte Anmeldetest...",
                'log_login_success': "[ERFOLG] Anmeldung erfolgreich! E-Mail und Passwort sind korrekt.",
                'log_login_fail_auth': "[FEHLER] Anmeldung fehlgeschlagen. Überprüfen Sie E-Mail oder Anwendungspasswort.",
                'log_login_fail_connection': "[FEHLER] Ein Verbindungsfehler oder ein anderer Fehler ist aufgetreten: {0}",
                'error_login_fields_main': "Füllen Sie die Felder E-Mail und Admin-Passwort aus.",
                'error_template_empty': "E-Mail-Betreff und -Inhalt dürfen nicht leer sein.",
                'error_no_file_selected': "Wählen Sie eine CSV-Datei aus.",
                'error_file_empty': "Die CSV-Datei ist leer.",
                'error_csv_header': "Ungültige Kopfzeilen in CSV.\nErwartet: {0}\nGefunden: {1}",
                'error_csv_read': "CSV-Datei konnte nicht gelesen werden: {0}",
                'error_participant_list_empty': "Teilnehmerliste ist leer. Fügen Sie jemanden hinzu.",
                'error_participant_list_short': "Es müssen mindestens 2 Teilnehmer vorhanden sein, um die Auslosung durchzuführen.",
                'log_draw_start': "--- Starte Auslosungsprozess ---",
                'log_btn_starting': "Auslosung & Senden...",
                'log_btn_idle': "Auslosung starten & E-Mails senden",
                'log_validating_participants': "Validiere {0} Teilnehmer...",
                'log_err_columns': "[FEHLER Datensatz {0}] Ungültige Spaltenanzahl. Überspringe.",
                'log_err_empty_row': "[FEHLER Datensatz {0}] Leerer 'Name' ('{1}') oder 'E-Mail' ('{2}'). Überspringe.",
                'log_err_invalid_email': "[FEHLER Datensatz {0}] Ungültige E-Mail: '{1}'. Überspringe.",
                'log_err_duplicate_name': "[FEHLER Datensatz {0}] Doppelter Name nach Normalisierung: '{1}' (von '{2}'). Name muss eindeutig sein! Abbruch.",
                'log_err_validation_short': "[FEHLER] Weniger als 2 Teilnehmer nach Validierung übrig. Abbruch.",
                'log_validation_success': "[INFO] {0} Teilnehmer erfolgreich validiert.",
                'log_draw_starting': "[INFO] Starte Auslosung...",
                'log_err_draw_failed': "[KRITISCHER FEHLER] Gültige Auslosung nach 1000 Versuchen nicht gefunden.\n[INFO] Möglicherweise unmöglich (z.B. 3 Personen in Paaren). Erneut versuchen oder Liste ändern.",
                'log_draw_success': "[ERFOLG] Auslosung erfolgreich abgeschlossen! Bereite E-Mail-Versand vor...",
                'log_show_results_title': "--- AUSLOSUNGSERGEBNISSE (Admin-Vorschau) ---",
                'log_show_results_pair': "  {0} -> {1}",
                'log_show_results_footer': "-----------------------------------------",
                'log_connecting_gmail': "Verbinde mit Gmail-Server...",
                'log_login_success_main': "[ERFOLG] Erfolgreich angemeldet als {0}",
                'log_login_fail_main': "[KRITISCHER FEHLER] Ungültige E-Mail oder Anwendungspasswort.",
                'log_login_fail_connection_main': "[KRITISCHER FEHLER] Anmeldefehler: {0}",
                'log_processing_send': "[Verarbeite] Sende an {0} ({1}) in {2}...",
                'log_err_template_key': "  [FEHLER] Fehler in der E-Mail-Vorlage. Fehlender Schlüssel: {0}. Verwenden Sie {{imie}} oder {{wylosowana_osoba}}.",
                'log_send_success': "  [ERFOLG] Information an {0} gesendet",
                'log_send_fail': "  [FEHLER] Senden an {0} fehlgeschlagen: {1}",
                'log_summary_title': "--- Versandzusammenfassung ---",
                'log_summary_success': "Informationen erfolgreich gesendet: {0}",
                'log_summary_errors': "Fehler: {0}",
                'log_err_critical': "[KRITISCHER FEHLER] Ein unerwarteter Fehler ist aufgetreten: {0}",
                'log_process_finished': "--- PROZESS ABGESCHLOSSEN ---",
                'log_generating_previews': "[INFO] Generiere Übersetzungsvorschau...",
                'log_previews_done': "[INFO] Vorschau fertig.",
                'log_err_googletrans_missing': "[FEHLER] 'googletrans' Bibliothek nicht installiert.\nFühren Sie 'pip install googletrans==4.0.0-rc1' im Terminal aus,\num die Übersetzung zu aktivieren."
            }
        }
        
        # Języki wspierane w interfejsie
        self.supported_ui_langs = {
            'pl': self.translations['pl']['lang_name'],
            'sk': self.translations['sk']['lang_name'],
            'en': self.translations['en']['lang_name'],
            'de': self.translations['de']['lang_name']
        }
        
        self.current_lang = 'pl'
        
        # --- ZMIANA: Dynamiczne mapy języków E-MAIL ---
        self.email_lang_map_display_to_code = {}
        self.email_lang_map_code_to_display = {}
        
        if GOOGLE_TRANS_AVAILABLE:
            for code, name in LANGUAGES.items():
                display_name = f"{name.capitalize()} ({code})"
                self.email_lang_map_display_to_code[display_name] = code
                self.email_lang_map_code_to_display[code] = display_name
        else:
            # Użyj tylko 4 języków UI jako fallback
            for code, name in self.supported_ui_langs.items():
                 self.email_lang_map_display_to_code[name] = code
                 self.email_lang_map_code_to_display[code] = name
        
        self.all_email_lang_list = sorted(list(self.email_lang_map_display_to_code.keys()))
        
        self.root.geometry("800x800")
        
        self.setup_styles()

        self.csv_path = tk.StringVar()
        self.log_queue = queue.Queue()
        self.preview_widgets = {} # Słownik na dynamiczne widgety podglądu

        self.create_widgets()
        self.check_log_queue()
        
        self.title_color_state = "GREEN" 
        self.animate_title()

    def t(self, key):
        return self.translations[self.current_lang].get(key, key)

    def normalize_text(self, text):
        text = text.lower()
        replacements = {
            'ę': 'e', 'ó': 'o', 'ą': 'a', 'ś': 's', 'ł': 'l', 'ż': 'z', 'ź': 'z', 'ć': 'c', 'ń': 'n'
        }
        for char, replacement in replacements.items():
            text = text.replace(char, replacement)
        return text

    def setup_styles(self):
        self.FESTIVE_GREEN = "#006400"
        self.FESTIVE_RED = "#B22222"
        BG_COLOR = "#FDFDFD"
        TEXT_COLOR = "#FFFFFF"

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('.', background=BG_COLOR, foreground='#333')
        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('TLabelframe', background=BG_COLOR, bordercolor=self.FESTIVE_GREEN)
        self.style.configure('TLabelframe.Label', background=BG_COLOR, foreground=self.FESTIVE_GREEN, font=('Helvetica', 12, 'bold'))
        self.style.configure('TNotebook.Tab', background=BG_COLOR, padding=[10, 5])
        self.style.map('TNotebook.Tab',
                       background=[('selected', self.FESTIVE_GREEN)],
                       foreground=[('selected', TEXT_COLOR)])
        self.style.configure('Success.TButton', background=self.FESTIVE_GREEN, foreground=TEXT_COLOR, font=('Helvetica', 10, 'bold'))
        self.style.map('Success.TButton', background=[('active', '#008000')])
        self.style.configure('Danger.TButton', background=self.FESTIVE_RED, foreground=TEXT_COLOR)
        self.style.map('Danger.TButton', background=[('active', '#CC0000')])
        self.style.configure('TButton', padding=5)
        self.style.configure('Treeview.Heading', background=self.FESTIVE_GREEN, foreground=TEXT_COLOR, font=('Helvetica', 10, 'bold'))
        self.style.map('Treeview', background=[('selected', '#B0E0B0')])
        self.style.map('TLabel', foreground=[('!disabled', '#000')])
        self.style.configure('Placeholder.TLabel', foreground='#555', font=('Helvetica', 9))
        self.style.configure('Instruction.TLabel', foreground='#333', font=('Helvetica', 10))

    def create_widgets(self):
        self.root.title(self.t('window_title'))
        
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.title_label = ttk.Label(
            main_frame,
            text=self.t('main_title'),
            font=("Helvetica", 18, "bold"),
            foreground=self.FESTIVE_GREEN,
            anchor="center"
        )
        self.title_label.pack(fill=tk.X, pady=(0, 15))

        self.main_notebook = ttk.Notebook(main_frame)
        self.main_notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0), side=tk.TOP) 

        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0), padx=5)

        self.footer_label = ttk.Label(
            footer_frame, 
            text=self.t('footer_author'), 
            foreground="#999", 
            anchor="e" 
        )
        self.footer_label.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        self.lang_combobox = ttk.Combobox(
            footer_frame, 
            values=list(self.supported_ui_langs.values()), 
            state='readonly', 
            width=20
        )
        self.lang_combobox.pack(side=tk.LEFT, padx=5)
        self.lang_combobox.current(0)
        self.lang_combobox.bind('<<ComboboxSelected>>', self.on_language_select)

        # --- Zakładka 1: Uczestnicy ---
        self.participants_tab_frame = ttk.Frame(self.main_notebook, padding="10")
        self.main_notebook.add(self.participants_tab_frame, text=self.t('tab1_participants'))

        self.input_frame = ttk.LabelFrame(self.participants_tab_frame, text=self.t('login_frame_title'), padding="10")
        self.input_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Dane logowania
        self.email_label = ttk.Label(self.input_frame, text=self.t('email_label'))
        self.email_label.grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.email_entry = ttk.Entry(self.input_frame, width=40)
        self.email_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        self.password_label = ttk.Label(self.input_frame, text=self.t('password_label'))
        self.password_label.grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.password_entry = ttk.Entry(self.input_frame, show="*", width=40)
        self.password_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        self.test_login_button = ttk.Button(self.input_frame, text=self.t('test_login_btn'), command=self.start_login_test)
        self.test_login_button.grid(row=0, column=2, rowspan=2, sticky="ns", padx=(10, 5), ipady=2)
        
        # Domyślny Język Maila
        self.default_lang_label = ttk.Label(self.input_frame, text=self.t('default_email_lang_label'))
        self.default_lang_label.grid(row=2, column=0, sticky="w", pady=5, padx=5)
        self.default_lang_combobox = ttk.Combobox(
            self.input_frame,
            values=self.all_email_lang_list,
            state='normal' # Umożliwia pisanie i autouzupełnianie
        )
        self.default_lang_combobox.grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        default_pl = self.email_lang_map_code_to_display.get('pl', self.all_email_lang_list[0])
        self.default_lang_combobox.set(default_pl)
        
        # Wewnętrzny Notebook (Uczestnicy)
        self.participants_notebook = ttk.Notebook(self.input_frame)
        self.participants_notebook.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(10, 5)) 
        self.input_frame.grid_rowconfigure(3, weight=1) 
        self.input_frame.columnconfigure(1, weight=1) 

        self.tab_manual = ttk.Frame(self.participants_notebook, padding="10")
        self.participants_notebook.add(self.tab_manual, text=self.t('participants_tab_manual'))
        self.create_manual_tab()

        self.tab_csv = ttk.Frame(self.participants_notebook, padding="10")
        self.participants_notebook.add(self.tab_csv, text=self.t('participants_tab_csv'))
        self.create_csv_tab() 

        # --- Zakładka 2: Treść Maila i Start ---
        self.template_tab_frame = ttk.Frame(self.main_notebook, padding="10")
        self.main_notebook.add(self.template_tab_frame, text=self.t('tab2_template'))

        self.template_notebook = ttk.Notebook(self.template_tab_frame)
        self.template_notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # Zakładka 2.1: Edytor
        self.editor_tab_frame = ttk.Frame(self.template_notebook, padding="10")
        self.template_notebook.add(self.editor_tab_frame, text=self.t('template_frame_title'))
        
        self.template_subject_label = ttk.Label(self.editor_tab_frame, text=self.t('template_subject_label'))
        self.template_subject_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_subject = ttk.Entry(self.editor_tab_frame)
        self.email_subject.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.email_subject.insert(0, self.t('default_subject'))
        
        self.template_body_label = ttk.Label(self.editor_tab_frame, text=self.t('template_body_label'))
        self.template_body_label.grid(row=1, column=0, sticky="nw", padx=5, pady=5)
        text_frame = ttk.Frame(self.editor_tab_frame)
        text_frame.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.email_body = tk.Text(text_frame, height=10, wrap=tk.WORD, undo=True, font=("Helvetica", 10)) 
        body_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.email_body.yview)
        self.email_body.configure(yscrollcommand=body_scroll.set)
        self.email_body.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        body_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.email_body.insert("1.0", self.t('default_body'))
        
        tag_button_frame = ttk.Frame(self.editor_tab_frame)
        tag_button_frame.grid(row=2, column=1, sticky="w", padx=5, pady=(0, 5))
        self.tags_label = ttk.Label(tag_button_frame, text=self.t('tags_label'))
        self.tags_label.pack(side=tk.LEFT, anchor='w')
        self.button_imie = ttk.Button(tag_button_frame, text=self.t('tag_btn_imie'), command=lambda: self.insert_tag_at_cursor("{imie}"))
        self.button_imie.pack(side=tk.LEFT, padx=(5, 2))
        self.button_wylosowana = ttk.Button(tag_button_frame, text=self.t('tag_btn_wylosowana'), command=lambda: self.insert_tag_at_cursor("{wylosowana_osoba}"))
        self.button_wylosowana.pack(side=tk.LEFT, padx=2)
        
        self.reset_template_button = ttk.Button(self.editor_tab_frame, text=self.t('reset_btn'), command=self.reset_template, style="Danger.TButton")
        self.reset_template_button.grid(row=0, column=3, rowspan=2, padx=10, sticky="ns")
        self.editor_tab_frame.columnconfigure(1, weight=1)

        # Zakładka 2.2: Podgląd i Edycja Tłumaczeń
        self.preview_tab_frame = ttk.Frame(self.template_notebook, padding="10")
        self.template_notebook.add(self.preview_tab_frame, text=self.t('preview_tab'))
        
        self.generate_preview_btn = ttk.Button(self.preview_tab_frame, text=self.t('generate_preview_btn'), command=self.generate_previews)
        self.generate_preview_btn.pack(fill=tk.X, pady=10)

        # Ta ramka będzie wypełniana dynamicznie
        self.dynamic_preview_frame = ttk.Frame(self.preview_tab_frame)
        self.dynamic_preview_frame.pack(fill=tk.BOTH, expand=True)

        # Sekcja Procesu (Process)
        self.process_frame = ttk.LabelFrame(self.template_tab_frame, text=self.t('start_frame_title'), padding="10")
        self.process_frame.pack(fill=tk.X, expand=False, pady=10)

        self.show_results_var = tk.BooleanVar(value=False)
        self.show_results_check = ttk.Checkbutton(
            self.process_frame, 
            text=self.t('show_results_check'), 
            variable=self.show_results_var
        )
        self.show_results_check.pack(pady=5, anchor='w')

        self.start_button = ttk.Button(self.process_frame, text=self.t('start_btn'), command=self.start_sending, style='Success.TButton')
        self.start_button.pack(fill=tk.X, ipady=10, pady=(10,0))

        # --- Zakładka 3: Raport (na żywo) ---
        self.output_frame = ttk.LabelFrame(self.main_notebook, text=self.t('report_frame_title'), padding="10")
        self.main_notebook.add(self.output_frame, text=self.t('tab3_report'))

        log_controls_frame = ttk.Frame(self.output_frame)
        log_controls_frame.pack(fill=tk.X, anchor='n')
        self.clear_log_button = ttk.Button(log_controls_frame, text=self.t('clear_log_btn'), command=self.clear_logs)
        self.clear_log_button.pack(side=tk.RIGHT, padx=0, pady=0)
        
        log_text_frame = ttk.Frame(self.output_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True) 

        self.log_text = tk.Text(log_text_frame, height=10, state="disabled", wrap=tk.WORD, bg="#FAFAFA", fg="#333")
        scrollbar = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def create_csv_tab(self):
        self.csv_file_label = ttk.Label(self.tab_csv, text=self.t('csv_file_label'))
        self.csv_file_label.grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.csv_label = ttk.Label(self.tab_csv, text=self.t('csv_no_file'))
        self.csv_label.grid(row=0, column=1, sticky="w", pady=5, padx=5)
        self.csv_button = ttk.Button(self.tab_csv, text=self.t('csv_select_btn'), command=self.select_csv_file)
        self.csv_button.grid(row=0, column=2, sticky="e", pady=5, padx=5)
        self.tab_csv.columnconfigure(1, weight=1)
        separator = ttk.Separator(self.tab_csv, orient='horizontal')
        separator.grid(row=1, column=0, columnspan=3, sticky='ew', pady=10, padx=5)
        
        self.csv_instruction_frame = ttk.LabelFrame(self.tab_csv, text=self.t('csv_instructions_title'), padding="10")
        self.csv_instruction_frame.grid(row=2, column=0, columnspan=3, sticky='ew', padx=5, pady=(5,0))
        
        self.csv_instruction_label = ttk.Label(
            self.csv_instruction_frame, 
            text=self.t('csv_instr_1'), 
            justify=tk.LEFT,
            style="Instruction.TLabel"
        )
        self.csv_instruction_label.pack(fill=tk.X, expand=True, padx=5, pady=5)

    def create_manual_tab(self):
        form_frame = ttk.Frame(self.tab_manual)
        form_frame.pack(fill=tk.X)
        self.manual_name_label = ttk.Label(form_frame, text=self.t('manual_name_label'))
        self.manual_name_label.grid(row=0, column=0, padx=5, sticky="w")
        self.manual_imie = ttk.Entry(form_frame)
        self.manual_imie.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.manual_partner_label = ttk.Label(form_frame, text=self.t('manual_partner_label'))
        self.manual_partner_label.grid(row=1, column=0, padx=5, sticky="w")
        self.manual_partner = ttk.Entry(form_frame)
        self.manual_partner.grid(row=1, column=1, padx=5, sticky="ew")
        
        self.manual_mail_label = ttk.Label(form_frame, text=self.t('manual_mail_label'))
        self.manual_mail_label.grid(row=2, column=0, padx=5, sticky="w")
        self.manual_mail = ttk.Entry(form_frame)
        self.manual_mail.grid(row=2, column=1, padx=5, sticky="ew")
        
        self.manual_lang_label = ttk.Label(form_frame, text=self.t('manual_lang_label'))
        self.manual_lang_label.grid(row=3, column=0, padx=5, sticky="w")
        
        self.manual_lang_combobox = ttk.Combobox(
            form_frame, 
            values=self.all_email_lang_list,
            state='normal'
        )
        self.manual_lang_combobox.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        
        self.add_button = ttk.Button(form_frame, text=self.t('add_to_list_btn'), command=self.add_to_tree)
        self.add_button.grid(row=1, column=2, rowspan=2, sticky="ns", padx=10)
        form_frame.columnconfigure(1, weight=1)
        
        self.list_frame = ttk.LabelFrame(self.tab_manual, text=self.t('participant_list_label'), padding="10")
        self.list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        cols = ('imię', 'partner', 'mail', 'jezyk')
        self.tree = ttk.Treeview(self.list_frame, columns=cols, show='headings', height=5)
        self.tree.heading('imię', text=self.t('tree_col_name'))
        self.tree.heading('partner', text=self.t('tree_col_partner'))
        self.tree.heading('mail', text=self.t('tree_col_mail'))
        self.tree.heading('jezyk', text=self.t('tree_col_lang'))
        self.tree.column('imię', width=120)
        self.tree.column('partner', width=120)
        self.tree.column('mail', width=150)
        self.tree.column('jezyk', width=100)
        
        tree_scroll = ttk.Scrollbar(self.list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        button_frame = ttk.Frame(self.list_frame)
        button_frame.pack(side=tk.RIGHT, anchor="n", padx=(5,0))
        self.edit_button = ttk.Button(button_frame, text=self.t('edit_btn'), command=self.edit_from_tree)
        self.edit_button.pack(fill=tk.X, pady=(0, 5))
        self.remove_button = ttk.Button(button_frame, text=self.t('delete_btn'), command=self.remove_from_tree, style='Danger.TButton')
        self.remove_button.pack(fill=tk.X)

    def on_language_select(self, event=None):
        selected = self.lang_combobox.get()
        for code, name in self.supported_ui_langs.items():
            if name == selected:
                self.switch_language(code)
                return

    def switch_language(self, lang):
        if lang == self.current_lang:
            return
        self.current_lang = lang
        
        # Aktualizuj wszystkie widgety
        self.root.title(self.t('window_title'))
        self.title_label.config(text=self.t('main_title'))
        self.footer_label.config(text=self.t('footer_author'))
        
        # Główne zakładki
        self.main_notebook.tab(self.participants_tab_frame, text=self.t('tab1_participants'))
        self.main_notebook.tab(self.template_tab_frame, text=self.t('tab2_template'))
        self.main_notebook.tab(self.output_frame, text=self.t('tab3_report'))

        # Zakładka 1
        self.input_frame.config(text=self.t('login_frame_title'))
        self.email_label.config(text=self.t('email_label'))
        self.password_label.config(text=self.t('password_label'))
        self.default_lang_label.config(text=self.t('default_email_lang_label'))
        self.test_login_button.config(text=self.t('test_login_btn'))
        self.participants_notebook.tab(self.tab_manual, text=self.t('participants_tab_manual'))
        self.participants_notebook.tab(self.tab_csv, text=self.t('participants_tab_csv'))
        
        # Zakładka 1 - Ręcznie
        self.manual_name_label.config(text=self.t('manual_name_label'))
        self.manual_partner_label.config(text=self.t('manual_partner_label'))
        self.manual_mail_label.config(text=self.t('manual_mail_label'))
        self.manual_lang_label.config(text=self.t('manual_lang_label'))
        self.add_button.config(text=self.t('add_to_list_btn'))
        self.list_frame.config(text=self.t('participant_list_label'))
        self.tree.heading('imię', text=self.t('tree_col_name'))
        self.tree.heading('partner', text=self.t('tree_col_partner'))
        self.tree.heading('mail', text=self.t('tree_col_mail'))
        self.tree.heading('jezyk', text=self.t('tree_col_lang'))
        self.edit_button.config(text=self.t('edit_btn'))
        self.remove_button.config(text=self.t('delete_btn'))

        # Zakładka 1 - CSV
        self.csv_file_label.config(text=self.t('csv_file_label'))
        self.csv_label.config(text=self.t('csv_no_file'))
        self.csv_button.config(text=self.t('csv_select_btn'))
        self.csv_instruction_frame.config(text=self.t('csv_instructions_title'))
        self.csv_instruction_label.config(text=self.t('csv_instr_1'))

        # Zakładka 2
        self.template_notebook.tab(self.editor_tab_frame, text=self.t('template_frame_title'))
        self.template_notebook.tab(self.preview_tab_frame, text=self.t('preview_tab'))
        self.template_subject_label.config(text=self.t('template_subject_label'))
        self.template_body_label.config(text=self.t('template_body_label'))
        self.tags_label.config(text=self.t('tags_label'))
        self.button_imie.config(text=self.t('tag_btn_imie'))
        self.button_wylosowana.config(text=self.t('tag_btn_wylosowana'))
        self.reset_template_button.config(text=self.t('reset_btn'))
        self.generate_preview_btn.config(text=self.t('generate_preview_btn'))
        self.process_frame.config(text=self.t('start_frame_title'))
        self.show_results_check.config(text=self.t('show_results_check'))
        self.start_button.config(text=self.t('start_btn'))

        # Zakładka 3
        self.output_frame.config(text=self.t('report_frame_title'))
        self.clear_log_button.config(text=self.t('clear_log_btn'))
        
        # Zaktualizuj domyślne wartości w polach
        current_subject = self.email_subject.get()
        current_body = self.email_body.get("1.0", tk.END).strip()
        
        all_default_subjects = [self.translations[lang]['default_subject'] for lang in self.translations]
        if current_subject in all_default_subjects:
            self.email_subject.delete(0, tk.END)
            self.email_subject.insert(0, self.t('default_subject'))
            
        all_default_bodies = [self.translations[lang]['default_body'].strip() for lang in self.translations]
        if current_body in all_default_bodies:
            self.email_body.delete("1.0", tk.END)
            self.email_body.insert("1.0", self.t('default_body'))

    def reset_template(self):
        self.main_notebook.select(2)
        if not messagebox.askyesno(self.t('confirm_title'), self.t('confirm_reset_template')):
            return
        self.email_subject.delete(0, tk.END)
        self.email_subject.insert(0, self.t('default_subject'))
        self.email_body.delete("1.0", tk.END)
        self.email_body.insert("1.0", self.t('default_body'))
        self.log(self.t('log_template_reset'))

    def add_to_tree(self):
        imie = self.manual_imie.get().strip()
        partner = self.manual_partner.get().strip()
        mail = self.manual_mail.get().strip()
        lang_display = self.manual_lang_combobox.get().strip()
        
        if not lang_display:
            lang_display = self.default_lang_combobox.get()
            
        lang_code = self.email_lang_map_display_to_code.get(lang_display)
        if not lang_code:
            found = False
            for key in self.all_email_lang_list:
                if lang_display.lower() in key.lower():
                    lang_display = key
                    lang_code = self.email_lang_map_display_to_code[key]
                    found = True
                    break
            if not found:
                lang_display = self.default_lang_combobox.get()
                lang_code = self.email_lang_map_display_to_code.get(lang_display, 'pl')
        
        if not imie or not mail:
            messagebox.showwarning(self.t('warn_title'), self.t('warn_empty_fields'))
            return
        if not is_valid_email(mail):
            messagebox.showwarning(self.t('warn_validation_title'), self.t('warn_invalid_email'))
            return
        
        self.tree.insert("", tk.END, values=(imie, partner, mail, lang_code))
        
        self.manual_imie.delete(0, tk.END)
        self.manual_partner.delete(0, tk.END)
        self.manual_mail.delete(0, tk.END)
        self.manual_lang_combobox.set('') 
        self.manual_imie.focus()

    def edit_from_tree(self):
        self.main_notebook.select(2)
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning(self.t('warn_no_selection_title'), self.t('warn_no_selection_edit'))
            return
        if len(selected_items) > 1:
            messagebox.showwarning(self.t('warn_no_selection_title'), self.t('warn_too_many_selected'))
            return
        item = selected_items[0]
        values = self.tree.item(item, 'values')

        self.manual_imie.delete(0, tk.END)
        self.manual_imie.insert(0, values[0])
        self.manual_partner.delete(0, tk.END)
        self.manual_partner.insert(0, values[1])
        self.manual_mail.delete(0, tk.END)
        self.manual_mail.insert(0, values[2])
        
        lang_code = values[3]
        lang_display = self.email_lang_map_code_to_display.get(lang_code, '')
        self.manual_lang_combobox.set(lang_display)

        self.tree.delete(item)
        self.manual_imie.focus()
        self.log(self.t('log_edit_loaded').format(values[0]))

    def remove_from_tree(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning(self.t('warn_no_selection_title'), self.t('warn_no_selection_delete'))
            return
        if messagebox.askyesno(self.t('confirm_title'), self.t('confirm_delete_records').format(len(selected_items))):
            for item in selected_items:
                self.tree.delete(item)

    def select_csv_file(self):
        self.main_notebook.select(2)
        path = filedialog.askopenfilename(
            title=self.t('csv_select_btn'),
            filetypes=[("Pliki CSV", "*.csv")]
        )
        if path:
            self.csv_path.set(path)
            self.csv_label.config(text=os.path.basename(path))
            self.log(self.t('log_file_selected').format(path))

    def log(self, message):
        self.log_queue.put(message)

    def insert_tag_at_cursor(self, tag):
        try:
            self.email_body.insert(tk.INSERT, tag)
            self.email_body.focus_set()
        except tk.TclError:
            self.log(self.t('log_insert_tag_error'))

    def check_log_queue(self):
        while not self.log_queue.empty():
            message = self.log_queue.get()
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        self.root.after(100, self.check_log_queue)

    def clear_logs(self):
        self.main_notebook.select(2)
        if not messagebox.askyesno(self.t('confirm_title'), self.t('confirm_clear_logs')):
            return
        while not self.log_queue.empty():
            try:
                self.log_queue.get_nowait()
            except queue.Empty:
                break
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")
        self.log(self.t('log_logs_cleared'))

    def start_login_test(self):
        self.main_notebook.select(2)
        sender_email = self.email_entry.get()
        app_password = self.password_entry.get()
        if not sender_email or not app_password:
            messagebox.showerror(self.t('error_title'), self.t('error_login_fields_empty'))
            return
        self.log(self.t('log_login_test_start'))
        self.test_login_button.config(state="disabled")
        threading.Thread(
            target=self.perform_login_test, 
            args=(sender_email, app_password),
            daemon=True
        ).start()

    def perform_login_test(self, sender_email, app_password):
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(GMAIL_SMTP_SERVER, GMAIL_SMTP_PORT, context=context) as server:
                server.login(sender_email, app_password)
                self.log(self.t('log_login_success'))
        except smtplib.SMTPAuthenticationError:
            self.log(self.t('log_login_fail_auth'))
        except Exception as e:
            self.log(self.t('log_login_fail_connection').format(e))
        finally:
            self.root.after(0, self.enable_test_button)

    def enable_test_button(self):
        self.test_login_button.config(state="normal")

    def start_sending(self):
        self.main_notebook.select(2)
        sender_email = self.email_entry.get()
        app_password = self.password_entry.get()
        if not sender_email or not app_password:
            messagebox.showerror(self.t('error_title'), self.t('error_login_fields_main'))
            return
        template_subject = self.email_subject.get()
        template_body = self.email_body.get("1.0", tk.END)
        if not template_subject or not template_body.strip():
            messagebox.showerror(self.t('error_title'), self.t('error_template_empty'))
            return
            
        show_results = self.show_results_var.get()

        participants_data = []
        selected_tab_index = self.participants_notebook.index(self.participants_notebook.select())
        
        if selected_tab_index == 0:
            for item in self.tree.get_children():
                participants_data.append(self.tree.item(item, 'values'))
        
        elif selected_tab_index == 1:
            csv_file_path = self.csv_path.get()
            if not csv_file_path:
                messagebox.showerror(self.t('error_title'), self.t('error_no_file_selected'))
                return
            try:
                with open(csv_file_path, mode='r', encoding='utf-8-sig') as file:
                    reader = list(csv.reader(file))
                    if not reader:
                        messagebox.showerror(self.t('error_title'), self.t('error_file_empty'))
                        return
                    header = [self.normalize_text(h.strip()) for h in reader[0]]
                    if header != OCZEKIWANE_NAGLOWKI:
                        messagebox.showerror(self.t('error_title'), self.t('error_csv_header').format(OCZEKIWANE_NAGLOWKI, header))
                        return
                    participants_data = reader[1:]
            except Exception as e:
                messagebox.showerror(self.t('error_title'), self.t('error_csv_read').format(e))
                return

        if not participants_data:
            messagebox.showerror(self.t('error_title'), self.t('error_participant_list_empty'))
            return
        if len(participants_data) < 2:
            messagebox.showerror(self.t('error_title'), self.t('error_participant_list_short'))
            return
        self.log(self.t('log_draw_start'))
        self.start_button.config(state="disabled", text=self.t('log_btn_starting'))
        
        threading.Thread(
            target=self.process_draw_and_send, 
            args=(sender_email, app_password, participants_data, template_subject, template_body, show_results),
            daemon=True
        ).start()

    def enable_start_button(self):
        self.start_button.config(text=self.t('start_btn'), state="normal")

    # --- ZMIANA: Dynamiczny podgląd ---
    def generate_previews(self):
        self.log(self.t('log_generating_previews'))
        self.main_notebook.select(2) # Przełącz na logi
        
        if not GOOGLE_TRANS_AVAILABLE:
            self.log(self.t('log_err_googletrans_missing'))
            messagebox.showwarning("Brak biblioteki", self.t('log_err_googletrans_missing'))
            return

        # 1. Zniszcz stare widgety podglądu
        for widget in self.dynamic_preview_frame.winfo_children():
            widget.destroy()
        self.preview_widgets = {} # Wyczyść słownik

        # 2. Zbierz unikalne języki z listy uczestników
        participants = [self.tree.item(item, 'values') for item in self.tree.get_children()]
        if not participants:
            self.log("[INFO] Lista uczestników jest pusta. Nie ma czego podglądać.")
            return

        unique_langs = set()
        for row in participants:
            if len(row) == 4 and row[3] and row[3] in self.email_lang_map_code_to_display:
                unique_langs.add(row[3])
        
        # 3. Usuń język źródłowy (polski)
        unique_langs.discard('pl')

        if not unique_langs:
            self.log("[INFO] Brak języków do tłumaczenia (wszyscy mają 'pl' lub nieprawidłowy kod).")
            return

        source_subject = self.email_subject.get()
        source_body = self.email_body.get("1.0", tk.END)

        # 4. Stwórz nowe, edytowalne widgety i uruchom tłumaczenie
        for lang_code in sorted(list(unique_langs)):
            lang_name = self.email_lang_map_code_to_display.get(lang_code, lang_code)
            
            lang_frame = ttk.Frame(self.dynamic_preview_frame, padding=5)
            lang_frame.pack(fill=tk.X, expand=True, pady=2)
            ttk.Label(lang_frame, text=f"{lang_name}:", font=('Helvetica', 10, 'bold')).pack(anchor='w')
            
            # Pole Tematu
            ttk.Label(lang_frame, text=self.t('template_subject_label')).pack(anchor='w', pady=(5,0))
            subject_widget = ttk.Entry(lang_frame, font=("Helvetica", 9))
            subject_widget.pack(fill=tk.X, expand=True, padx=5, pady=2)
            
            # Pole Treści
            ttk.Label(lang_frame, text=self.t('template_body_label')).pack(anchor='w', pady=(5,0))
            body_widget = tk.Text(lang_frame, height=4, wrap=tk.WORD, bg='#FFF', font=("Helvetica", 9))
            body_widget.pack(fill=tk.X, expand=True, padx=5, pady=2)

            self.preview_widgets[lang_code] = {
                'subject': subject_widget,
                'body': body_widget
            }

            self.log(f"[INFO] Tłumaczenie na {lang_code}...")
            threading.Thread(
                target=self.translate_and_update_widget,
                args=(source_subject, source_body, lang_code, subject_widget, body_widget),
                daemon=True
            ).start()
        
        self.log(self.t('log_previews_done'))


    def translate_and_update_widget(self, subject, body, lang_code, subject_widget, body_widget):
        """Funkcja pomocnicza dla wątku tłumaczenia podglądu"""
        translated_subject = get_translation(subject, lang_code, source_lang='pl')
        translated_body = get_translation(body, lang_code, source_lang='pl')
        
        self.root.after(0, self.update_preview_text, subject_widget, body_widget, translated_subject, translated_body)
        
    def update_preview_text(self, subject_widget, body_widget, subject_content, body_content):
        """Bezpiecznie aktualizuje edytowalne widgety z wątku głównego"""
        subject_widget.config(state='normal')
        subject_widget.delete(0, tk.END)
        subject_widget.insert(0, subject_content)
        
        body_widget.config(state='normal')
        body_widget.delete("1.0", tk.END)
        body_widget.insert("1.0", body_content)


    def process_draw_and_send(self, sender_email, app_password, participants_data, template_subject, template_body, show_results):
        participants = []
        try:
            self.log(self.t('log_validating_participants').format(len(participants_data)))
            imiona_uczestnikow = set()
            
            default_lang_display = self.default_lang_combobox.get()
            default_lang_code = self.email_lang_map_display_to_code.get(default_lang_display, 'pl')
            
            for i, row in enumerate(participants_data):
                if len(row) != 4:
                    self.log(self.t('log_err_columns').format(i+1))
                    continue
                
                imie = str(row[0]).strip()
                partner = str(row[1]).strip()
                mail = str(row[2]).strip()
                jezyk = str(row[3]).strip().lower()
                
                if not jezyk or jezyk not in self.email_lang_map_code_to_display:
                    jezyk = default_lang_code 
                
                imie_norm = self.normalize_text(imie)
                partner_norm = self.normalize_text(partner)
                
                if not imie or not mail:
                    self.log(self.t('log_err_empty_row').format(i+1, imie, mail))
                    continue
                if not is_valid_email(mail):
                    self.log(self.t('log_err_invalid_email').format(i+1, mail))
                    continue
                if imie_norm in imiona_uczestnikow:
                    self.log(self.t('log_err_duplicate_name').format(i+1, imie_norm, imie))
                    return
                imiona_uczestnikow.add(imie_norm)
                
                participants.append({
                    'name': imie, 
                    'partner': partner, 
                    'mail': mail,
                    'lang': jezyk,
                    'name_norm': imie_norm,
                    'partner_norm': partner_norm
                })
                
            if len(participants) < 2:
                self.log(self.t('log_err_validation_short'))
                return
            self.log(self.t('log_validation_success').format(len(participants)))
            self.log(self.t('log_draw_starting'))
            draw_results = self.perform_draw(participants)
            if draw_results is None:
                self.log(self.t('log_err_draw_failed'))
                return
            
            self.log(self.t('log_draw_success'))

            if show_results:
                self.log(self.t('log_show_results_title'))
                name_map = {p['name_norm']: p['name'] for p in participants}
                for giver_norm, receiver_norm in draw_results.items():
                    giver_display = name_map.get(giver_norm, giver_norm)
                    receiver_display = name_map.get(receiver_norm, receiver_norm)
                    self.log(self.t('log_show_results_pair').format(giver_display, receiver_display))
                self.log(self.t('log_show_results_footer'))

            if not GOOGLE_TRANS_AVAILABLE:
                self.log(self.t('log_err_googletrans_missing'))
                messagebox.showwarning("Brak biblioteki", self.t('log_err_googletrans_missing'))

            # --- ZMIANA: Pobierz ręczne tłumaczenia ---
            translation_overrides = {}
            for lang_code, widgets in self.preview_widgets.items():
                subject = widgets['subject'].get()
                body = widgets['body'].get("1.0", tk.END)
                translation_overrides[lang_code] = {'subject': subject, 'body': body}
            # --- Koniec zmiany ---

            self.log(self.t('log_connecting_gmail'))
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(GMAIL_SMTP_SERVER, GMAIL_SMTP_PORT, context=context) as server:
                try:
                    server.login(sender_email, app_password)
                    self.log(self.t('log_login_success_main').format(sender_email))
                except smtplib.SMTPAuthenticationError:
                    self.log(self.t('log_login_fail_main'))
                    return
                except Exception as e:
                    self.log(self.t('log_login_fail_connection_main').format(e))
                    return
                
                sukces_count = 0
                blad_count = 0
                participant_map = {p['name_norm']: p for p in participants}
                
                for giver_name_norm, receiver_name_norm in draw_results.items():
                    giver_obj = participant_map[giver_name_norm]
                    receiver_obj = participant_map[receiver_name_norm]
                    
                    giver_name_oryginal = giver_obj['name']
                    giver_mail = giver_obj['mail']
                    giver_lang = giver_obj['lang'] 
                    receiver_name_oryginal = receiver_obj['name']
                    
                    self.log(self.t('log_processing_send').format(giver_name_oryginal, giver_mail, giver_lang))
                    
                    try:
                        # --- ZMIANA: Sprawdź ręczne poprawki ---
                        if giver_lang in translation_overrides:
                            translated_subject = translation_overrides[giver_lang]['subject']
                            translated_body = translation_overrides[giver_lang]['body']
                            self.log(f"  [INFO] Używam ręcznie edytowanego tłumaczenia dla {giver_lang}.")
                        else:
                            translated_subject = get_translation(template_subject, giver_lang, source_lang='pl')
                            translated_body = get_translation(template_body, giver_lang, source_lang='pl')
                        # --- Koniec zmiany ---
                        
                        subject = translated_subject.format(
                            imie=giver_name_oryginal, 
                            wylosowana_osoba=receiver_name_oryginal
                        )
                        body = translated_body.format(
                            imie=giver_name_oryginal, 
                            wylosowana_osoba=receiver_name_oryginal
                        )
                    except KeyError as e:
                        self.log(self.t('log_err_template_key').format(e))
                        blad_count += 1
                        continue
                    except Exception as e:
                         self.log(f"  [BŁĄD] Inny błąd tłumaczenia/formatowania: {e}")
                         blad_count += 1
                         continue

                    msg = MIMEMultipart()
                    msg['From'] = f"Tajemniczy Mikołaj <{sender_email}>"
                    msg['To'] = giver_mail
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain', 'utf-8'))
                    
                    try:
                        server.sendmail(sender_email, giver_mail, msg.as_string())
                        self.log(self.t('log_send_success').format(giver_name_oryginal))
                        sukces_count += 1
                    except Exception as e:
                        self.log(self.t('log_send_fail').format(giver_name_oryginal, e))
                        blad_count += 1
                
                self.log(self.t('log_summary_title'))
                self.log(self.t('log_summary_success').format(sukces_count))
                self.log(self.t('log_summary_errors').format(blad_count))
        except Exception as e:
            self.log(self.t('log_err_critical').format(e))
        finally:
            self.root.after(0, self.enable_start_button)
            self.log(self.t('log_process_finished'))

    def perform_draw(self, participants):
        MAX_TRIES = 1000
        givers = [p['name_norm'] for p in participants]
        receivers = list(givers)
        participant_map = {p['name_norm']: p for p in participants} 
        for _ in range(MAX_TRIES):
            random.shuffle(receivers)
            is_valid_draw = True
            draw = {}
            for i in range(len(givers)):
                giver_name_norm = givers[i]
                receiver_name_norm = receivers[i]
                if giver_name_norm == receiver_name_norm:
                    is_valid_draw = False
                    break
                giver_partner_norm = participant_map[giver_name_norm]['partner_norm']
                if giver_partner_norm and giver_partner_norm == receiver_name_norm:
                    is_valid_draw = False
                    break
                draw[giver_name_norm] = receiver_name_norm
            if is_valid_draw:
                return draw 
        return None 

    def animate_title(self):
        try:
            if self.title_color_state == "GREEN":
                self.title_label.config(foreground=self.FESTIVE_RED)
                self.title_color_state = "RED"
            else:
                self.title_label.config(foreground=self.FESTIVE_GREEN)
                self.title_color_state = "GREEN"
            self.root.after(1200, self.animate_title)
        except tk.TclError:
            pass

# --- Uruchomienie aplikacji ---
if __name__ == "__main__":
    if not GOOGLE_TRANS_AVAILABLE:
        print("--- OSTRZEŻENIE ---")
        print("Nie znaleziono biblioteki 'googletrans'. Funkcja tłumaczenia będzie wyłączona.")
        print("Dostępne będą tylko języki: PL, SK, EN, DE.")
        print("Aby włączyć pełne tłumaczenie i listę wszystkich języków, zainstaluj ją w terminalu:")
        print("pip install googletrans==4.0.0-rc1")
        print("-------------------")
        
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
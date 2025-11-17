# Generator do tajnego losowania osoby, której ktoś ma dać prezent na Wigilię

"Projekt" stworzony na potrzeby Wigilii 2024, aby zdalnie przeprowadzić losowanie wśród osób dorosłych. Kod wygenerowany przez GPT 4o

Zasady:
* nie można wylosować siebie
* nie można wylosować swojego partnera
* każda osoba może być wy

## Podjęte kroki
1. Napisałem prompt do GPT 4o wytyczający kryteria.
2. Wybrałem konto gmailowe do skryptowego rozesłania wiadomości mailowych
3. Na wybranym koncie gmail włączyłem weryfikację dwuetapową (wymagane by użyć biblioteki Python - smtplib, która wysyła maile)
4. Wygenerowałem hasło aplikacji - 16 znakowe, które daje odpowiednie uprawnienia skryptowi (jeśli poda się je wraz z odpowiednim loginem oczywiście).

## Na co trzeba uważać:
* użyj nieistotnej skrzynki odbiorczej do takich celów
* nie podawaj swojego hasła! Podaj hasło aplikacji. Do tego dobra jest skrzynka gmail.
* unikaj literówek - łatwo sie pomylić przy podawaniu aresów mail, a jak już do tego dojdzie - trzeba powtarzać losowanie.
  

## Spostrzeżenia:

* Pomysł został zrealizowany w 15-20 minut - stało się to dzięki GPT.
* Poza jednym błędem ludzkim wszystko przeszło poprawnie.
* Całość pozwala zwątpić w przyszłość zawodu programisty jaki to zawód jest w obecnym kształcie.


# Mobi---PTP
Program automatyzuje pracę w obszarze operacyjnym działu PTP

Autor: Kamila Dudzińska

Projekt: Program 'Mobi' do automatyzacji maili 
	 dedykowany dla procesów operacyjnych dla działu zakupów (Procurement)

Źródło: sama wygenerowałam próbkę danych na potrzeby spr. automatyzacji

Cel: Stworzenie programu do analizy tabeli excel z danymi o zamówieniach w systemie SAP ARIBA oraz automatycznego wysyłania maili do kupców z przypomnieniem. Program generuje też raport dla administratora, do kogo maile zostały wysłane i jakie są statystyki zamówień. Dzięki temu można jednym kliknięciem zaoszczędzić sporo FTE, a administrator może szybko uzyskać realny "stan rzeczy".


Jak działa program:
1. Program iteruje wiersz po wierszu w tabeli za zamówieniami. 
2.Jeśli znajdzie zamówienie (PO) ze statusem "ordered" ("zamówione") to sprawdzi dodatkowo prognozowaną datę dostawy(delivery date).
2. Jeżeli data dostawy jest w przeszłości (dzisiaj odjąć 3 dni*) to potraktuje to jako informację do wykonania zadania --> wyśle maila z przypomnieniem o zrobieniu GR do kupca.
3. Program czyta dane z tabeli excel, jeżeli jeden kupiec będzie miał kilka różnych zamówień, to zostaną wysłane do niego szczegóły o wszystkich zamówieniach
4. Po wykonaniu zadania program poinformuje administratora, gdzie udało mu się wysłać maila - w przypadku aktywnej konsoli IDE oraz dodatkowo wyśle raport ze statystykami w formacie pdf na maila administratora. 

Zalety projektu:
--> odpowiada na realny problem w wielu procesach operacyjnych, gdzie wymagane jest sprawdzanie i repetetywne wysyłanie przypominajek/follow-upów
--> zmniejsza problem z tworzeniem przyjęcia GR przez nietechnicznych kupców, którzy często nie pilnują swoich zamówień tłumacząc to jako - "W Aribie Guided Buing nie da się filtrować po statusach". 
--> dzięki monitorowaniu stanu zamówień (PO) i przyjęć (GR) można zmniejszyć "invoice overdue" (niepłacenie faktur na czas) a co za tym idzie - zminimalizować ryzyko kłopotów z dostawcami, czy utraty wizerunku
--> administrator programu otrzymuje statystyki, dzięki czemu łatwiej kontrolować proces GR
--> program automatyzuje pracę w obrębie działu zakupów
--> program napisany pod typowe środowisko korporacyjne z zalogowanym "Outlookiem"
--> program dedykowany SAP ARIBA (z  racji, że pracuję na tym programie jako key user) ale można go szybko dopasować do innych systemów - wystarczy przeanalizować raporty generowane np. przez SAP MM, czy inny dowolny program.


Tabela z zamówieniami (na żółto te zamówienia, gdzie Mobi powienien wysłać przypominajkę):
<img width="1508" height="198" alt="image" src="https://github.com/user-attachments/assets/271c8ce8-e929-41fd-82a0-b3b7079140cf" />

Fragment kodu: Czyszczenie danych
<img width="763" height="275" alt="czyszczenie" src="https://github.com/user-attachments/assets/89331507-b596-449a-b70c-d637b55a88f2" />

Fragment kody: połączenie z outlookiem, pętla oraz instrukcje warunkowe:
<img width="887" height="810" alt="image" src="https://github.com/user-attachments/assets/e8838fd7-9582-40da-a8f9-a573813e21a5" />

Email administratora:
<img width="776" height="496" alt="image" src="https://github.com/user-attachments/assets/cfa1b47f-89cf-49aa-a0c5-0e09831deb8c" />

Statystyki z załącznika:
<img width="577" height="391" alt="image" src="https://github.com/user-attachments/assets/50778feb-eb99-49d1-9fcf-39030664d346" />



Kod: cały kod znajduje się w osobnym pliku

Przykładowe fragmenty kodu oraz screen z maila i raportów.

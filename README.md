# PretvoriKverijaMakroaVoVBA
Претвара листа на кверија во текст фајлови во VBA *.bas фајл кој потоа може да се користи во Акцес или Ексел како таков

Се извезуваат кверијата од Акцес фајл во некоја папка со користење на Utils.ExportQuerySQL(pateka). 
Потоа се местат папките и табелите за тоа каде да ги најде кверијата и кои имиња на табели да ги замени со кои, 
наведени во тагот iminjaTabeli одвоени со „цевка“ (|) во PretvoriKverijaVoVBA.exe.config

Функцијата Utils.IzveziKverijaIMakroaVoModuli(pateka) е експериментална. Ги извезува кверијата во соодветни папки со 
името на макрото во кое кверијата се повикуваат, во Акцес фајлот. Ова би значело дека во секоја од овие папки ќе треба 
да се мести нова конфигурација на претворачот, но во случај едно макро да има 30+ кверија, помага во прегледност.  
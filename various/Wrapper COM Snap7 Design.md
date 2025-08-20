# SCOPO DI QUESTA LISTA DI DESIGN, REQUISITI, FUNZIONALITA

**Considerando** una asplicazione **industriale** di questo tipo: **PLC S7 Siemens con protocollo S7** ethernet che comunica con un **PC Windows (dev/exec)**, dove un applicativo **legacy VB6 SP6** con String unicode, Decimal, currency, tipo dati nativi, utilizzando un **WRAPPER COM VB.NET** con **SNAP7**, accede sia in lettura sia in scrittura ai dati presenti nel PLC.

Definire il design i requisiti e le funzionalità del **Wrapper COM Snap7** da realizzare in **VB.NET**, utilizzando lambiente di sviluppo **Visual studio Comunity edition 2022** e il  .**NET Framework 4.8**.

**Requisito primario** la **compatibilità totale** del **Wrapper COM** con **VB6 SP6**, per l'utilizzo sia **Early-binding che** che **Late-Binding**.

## Mappatura tipi dati da PLC S7 al WRAPPER (VB.NET + Snap7) a COM a VB6 (lingua: ITA)

Di seguito trovi una tabella operativa e completa per il mapping dei tipi S7 più usati, con limiti, dimensioni, formato in memoria e i problemi pratici lato VB6 + soluzioni consigliate. Fonti: documentazione SIEMENS, Microsoft (VB.NET / COM / BSTR / Variant) e repo Snap7 (davenardella). ([Automazione Integrata Totale][1], [Microsoft Learn][2], [GitHub][3])

> Nota veloce: Snap7 espone i dati come buffer byte; il wrapper VB.NET deve **decodificare** in base al tipo S7 (endianness, header STRING, BCD per DATE\_AND\_TIME). Snap7/S7 usano per default ordine *big-endian* per variabili multi-byte (con istruzioni per little-endian quando necessario). ([snap7.sourceforge.net][4], [Automazione Integrata Totale][5])

---

## Tabella di mapping (rilevante / operativa)

| PLC (tipo / limiti / dimensione / intervallo / precisione / formato memoria)                                         |                                                  VB.NET / Snap7 (tipo, limiti) | Tecn. COM (tipo, limiti)                    | VB6 SP6 (tipo, limiti)                                                                              | Problemi pratici in VB6 (overflow, mapping, endianess, note)                                                                                                                                                                                                                                                                                                     |
| -------------------------------------------------------------------------------------------------------------------- | -----------------------------------------------------------------------------: | ------------------------------------------- | --------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **BOOL** - 1 bit (occupato dentro byte), {0,1} - memorizzazione bitwise                                              |                                        `Boolean` (System.Boolean) - True/False | `VT_BOOL` / VARIANT (VT\_BOOL)              | **Boolean** (VB6) (True/False)                                                                      | Allineamento: SNAP7 restituisce byte - il wrapper deve estrarre il bit corretto; attenzione a indirizzi %X / bit offset. (Siemens bit addressing). ([Siemens][6])                                                                                                                                                                                                |
| **SINT (INT8)** - 1 byte signed `-128..127`                                                                          |                                                         `SByte` (System.SByte) | `VT_I1` / custom marshalling                | **Byte/Integer** - VB6 non ha SByte; usare `Byte` (0..255) o `Integer` con casting                  | Negativo: VB6 non ha SByte nativo; usare `CShort`/`Integer` per segnato oppure mappare in `Integer` e interpretare due's complement. Endian: n/a.                                                                                                                                                                                                                |
| **USINT / BYTE** - 1 byte unsigned `0..255`                                                                          |                                                           `Byte` (System.Byte) | `VT_UI1` / VARIANT                          | **Byte** (VB6: 0..255)                                                                              | OK se wrapper restituisce 1 byte; attenzione a String(binaries).                                                                                                                                                                                                                                                                                                 |
| **INT (INT16)** - 2 bytes signed `-32,768..32,767` (word-aligned)                                                    |                                                              `Short` / `Int16` | `VT_I2`                                     | **Integer (VB6)** *NB:* VB6 *Integer* - 16-bit.                                                     | Fine se mappari Short-VB6 Integer. Endian: S7 big-endian - Snap7/helper deve ordinare correttamente. ([Siemens][6])                                                                                                                                                                                                                                              |
| **UINT / WORD** - 2 bytes unsigned `0..65535`                                                                        |                                                   `UShort` / `UInt16` (VB.NET) | `VT_UI2` / VARIANT                          | **Integer (VB6)** (signed 16-bit) - non esiste UInteger 16-bit                                      | Possibile overflow >32767; soluzione: mappare a **Long (VB6)** o Variant numerico, oppure usare `CLng`/`CInt` con controllo range.                                                                                                                                                                                                                               |
| **DINT (INT32)** - 4 bytes signed `-2,147,483,648..2,147,483,647`                                                    |                                                            `Integer` / `Int32` | `VT_I4`                                     | **Long (VB6)** (32-bit)                                                                             | OK diretto. Endian: attenzione all'ordine delle parole quando DB ottimizzati; usare helper Snap7/BitConverter con ordine corretto. ([GitHub][3])                                                                                                                                                                                                                 |
| **UDINT / DWORD** - 4 bytes unsigned `0..4,294,967,295`                                                              |                                                          `UInteger` / `UInt32` | `VT_UI4`                                    | **Long (VB6)** - signed 32-bit - non può rappresentare >2,147,483,647                               | Rischio overflow per valori >2^31-1. Soluzioni: trasferire come `Double`/`Variant` o come STRING; oppure usare controllo e mappare in `Double` o `Currency` (se valori interi e rientrano).                                                                                                                                                                      |
| **LINT (INT64)** - 8 bytes signed                                                                                    |                                                               `Long` / `Int64` | `VT_I8` *(Automation historically limited)* | **VB6 non ha Int64 nativo** (solo Long = 32-bit); su 64-bit VBA esiste LongLong, ma VB6 classico NO | Problema critico: COM Automation / VB6 non supporta VT\_I8 in standard Variant; soluzioni: 1) WMARSHAL nel wrapper e passare come `Decimal`/`Variant` (CDec) o `Currency` con scala, 2) passare come stringa, 3) spezzare in due Long. Raccomandazione: usare `Variant` con `Decimal` subtype o BSTR (string) per sicurezza. ([Microsoft Learn][7], [GitHub][3]) |
| **REAL (FLOAT32)** - 4 bytes IEEE754                                                                                 |                                                       `Single` (System.Single) | `VT_R4`                                     | **Single (VB6)**                                                                                    | OK diretto; attenzione rappresentazione IEEE e endian; use BitConverter con ordine.                                                                                                                                                                                                                                                                              |
| **LREAL (FLOAT64)** - 8 bytes IEEE754                                                                                |                                                                       `Double` | `VT_R8`                                     | **Double (VB6)**                                                                                    | OK diretto.                                                                                                                                                                                                                                                                                                                                                      |
| **CHAR** - 1 byte ASCII (S7 `CHAR`)                                                                                  |                           `Byte` / convertire via `Encoding.ASCII.GetString()` | `VT_UI1` o BSTR dopo conversione            | **String (BSTR) o Byte**                                                                            | Map su VB6 `String` richiede conversione da byte-Unicode (BSTR) e rimozione eventuale `0` terminator; CHAR singolo - usare `Chr` / CStr.                                                                                                                                                                                                                         |
| **STRING\[n] (S7)** - stored con header (2 byte: maxLen + curLen) + data (max 254 default) - fino a 254 char (ASCII) | In VB.NET: `String` (System.String) dopo parsing del buffer (rimuovere header) | COM: `BSTR` (Unicode)                       | **String (VB6)** (BSTR, Unicode)                                                                    | **Importantissimo**: S7 STRING ha header (maxLen+len) - NON mandare raw buffer a BSTR. Wrapper deve: leggere header, estrarre dati, convertire encoding ASCII-Unicode e restituire BSTR. Overflow: lunghezze >VB6 limiti; tagliare o gestire. ([Siemens Support][8], [GitHub][3])                                                                                |
| **WCHAR / WSTRING** - WCHAR 2 byte, WSTRING wide chars                                                               |                                                       `Char`/`String` (UTF-16) | `BSTR`                                      | **String (VB6)**                                                                                    | Se PLC espone WString (2-byte) si può marshallare direttamente su BSTR se wrapper mantiene UTF-16. Attenzione endianness. ([Siemens Support][9])                                                                                                                                                                                                                 |
| **DATE** (S7 DATE) - 2 bytes (days since 1990-01-01)                                                                 |                                         VB.NET `DateTime` (convertire da days) | COM `DATE` / `VT_DATE`                      | **Date (VB6)** (8 bytes double)                                                                     | Trasformazione necessaria: convertire offset (days) - DateTime; wrapper deve calcolare. ([Automazione Integrata Totale][1])                                                                                                                                                                                                                                      |
| **TIME** / **TOD** - 4 bytes (ms since midnight) / TIME\_OF\_DAY                                                     |                                                 VB.NET `TimeSpan` / `DateTime` | COM `VT_DATE` (o double)                    | **Date** o `Long` ms                                                                                | Convertire millisecondi - TimeSpan/DateTime.                                                                                                                                                                                                                                                                                                                     |
| **DATE\_AND\_TIME (DT)** - 8 bytes BCD (year..ms) - formato BCD/DT S7                                                |                                              VB.NET `DateTime` dopo BCD decode | COM `DATE` / `VT_DATE`                      | **Date** (VB6)                                                                                      | Necessita decodifica BCD dal buffer S7 - creare DateTime; Snap7 non fa automaticamente la conversione "semantica". ([Siemens Support][10])                                                                                                                                                                                                                       |
| **TIME\_SPAN / LWORD / 64-bit unsigned** - 8 bytes unsigned                                                          |                                                             `UInt64` / `ULong` | `VT_UI8` (Automation limiti)                | **NO native 64-bit unsigned**                                                                       | Vedi LINT: usare Variant/Decimal o stringa; attenzione overflow.                                                                                                                                                                                                                                                                                                 |
| **Bit arrays / DB strutturati / UDT** - vari dimensionamenti                                                         |                              In VB.NET: `byte[]`, strutture `Struct` + Marshal | COM: SAFEARRAY / custom COM object          | VB6: `Variant` array o SAFEARRAY via COM                                                            | Problemi: layout dei struct (packing/word boundaries/optimized DB) - usare descrizione esplicita DB o esportare campi semplici. Raccomandazione: usare DB con elementi scalari o serializzare in JSON/BSTR. ([Siemens][6])                                                                                                                                       |

---

## Indici (veloce, per ricerca)

### Indice Tipi dato PLC (presenti nella tabella)

BOOL, SINT, USINT/BYTE, INT, UINT/WORD, DINT, UDINT/DWORD, LINT, UDINT64/LWORD, REAL, LREAL, CHAR, STRING, WCHAR/WSTRING, DATE, TIME, TIME\_OF\_DAY, DATE\_AND\_TIME (DT), UDT/STRUCT/DB, BIT ARRAY.

### Indice Tipi dato VB.NET (presenti nella tabella)

Boolean, SByte, Byte, Short(Int16), UShort(UInt16), Integer(Int32), UInteger(UInt32), Long(Int64), ULong(UInt64), Single, Double, Decimal, String, DateTime, TimeSpan, Byte\[].

### Indice Tipi dato VB6 (presenti nella tabella)

Boolean, Byte, Integer (16-bit), Long (32-bit), Single, Double, String (BSTR), Date, Variant (incl. Decimal subtype), Currency.

---

## Info tecnologie impiegate - dettagliata (sintesi esecutiva)

* **Siemens S7 (S7-300/400/1200/1500)**: tipi PLC definiti in STEP7 / TIA Portal; STRING ha header (2 byte), DATE\_AND\_TIME - BCD su 8 byte; memoria a word-boundaries; default multi-byte = *big-endian*, con istruzioni per little-endian se richiesto. ([Siemens Support][8], [Automazione Integrata Totale][1])
* **Snap7 (davenardella)**: client/library C++ che espone API per leggere/write memory (DB, M, I, Q) come buffer di byte; include DLL precompilate per Win32/Win64 e utility; - *runtime endian-aware* (ma il wrapper deve comunque interpretare i buffer). Consigli: usare le funzioni helper (o Sharp7 wrapper) per Get/Set di tipi o gestire `BitConverter` + ordine byte. ([GitHub][3], [snap7.sourceforge.net][4])
* **WRAPPER COM in VB.NET**: layer che usa Snap7 native DLL (P/Invoke) o wrapper .NET (Sharp7), espone interfaccia COM (Registrare assembly con `RegAsm` oppure creare COM-visible classi). Marshaling: stringhe gestite come BSTR, numerici come VT\_I4/VT\_R8; per tipi 64-bit e unsigned attenzione a compatibilità Automation. Per stringhe S7: wrapper deve rimuovere header e convertire encoding. ([Microsoft Learn][11], [GitHub][3])
* **VB6 legacy**: cliente COM che riceve BSTR / Variants / SAFEARRAY; VB6 non ha molti tipi moderni (p.es. no Int64 nativo); Decimal - usabile solo come **subtype di Variant** (`CDec`) - quindi wrapper deve adattarsi o passare BSTR. ([Microsoft Learn][7])

---

## Hardware necessario (pratico, lista esecutiva)

* PLC Siemens (es. S7-1200 / S7-1500 o S7-300/400 a seconda impianto). Documentazione PLC TIA Portal per definizione DB. ([Siemens Industry Cache][12])
* Rete Ethernet industriale: switch Gigabit, cavi CAT5e/6, isolamento se richiesto.
* PC Windows (dev/exec) con: .NET Framework / .NET runtime compatibile, Visual Studio per wrapper VB.NET, VB6 runtime (o IDE per debug) - se legacy app su VB6 SP6.
* Snap7 DLLs (win32/win64) dal repo ufficiale (in /build/bin). Registrare e linkare al wrapper. ([GitHub][3])
5. Strumenti di test: HMITracer / clientdemo (in snap7.utility) per debug rete e pacchetti. ([GitHub][3])

---

## Note operative dettagliate (passo-passo high-level)

* **Definizione DB in PLC**: evitare UDT/composti non necessari per comunicazione. Prediligi DB lineari con elementi scalari (INT, DINT, REAL, STRING\[n]). (Siemens guideline). ([Siemens][6])
* **Progettare wrapper VB.NET**:

   * Usare Snap7 client DLL (o Sharp7) per read/write DB come `byte[]`. ([GitHub][3])
   * Per ogni elemento: eseguire parsing secondo tipo S7 (es. per STRING leggere 2 byte header - lunghezza - substring). ([Siemens Support][8])
   * Convertire endian se necessario (BitConverter + Reverse se Snap7 non ha helper). Snap7 - endian-aware ma fai test end-2-end. ([snap7.sourceforge.net][4])
   * Esporre API COM-Visible: metodi `GetTagAsString`, `GetTagAsLong`, `GetTagAsDate` che ritornano BSTR/Variant compatibili VB6. Usare `[ComVisible(true)]`, `ClassInterfaceType.None`, `Guid(...)`. Registrare con `RegAsm`.
* **Regole di marshaling per VB6**: preferire ritornare: BSTR per testi, `Double`/`Long` per numerici entro range VB6; per 64-bit passare BSTR o Variant(Decimal). ([Microsoft Learn][13])
* **Test**: testare valori estremi (limiti, negative, strings con 0x00), test di endian swapping e DB ottimizzati (in S7 i DB ottimizzati possono influire su layout). ([Siemens][6])

---

## Note aggiuntive dettagliate (rischi & mitigazioni)

* **DB ottimizzati / packing**: possono modificare offset degli elementi: usa simbolic access (nomi dei tags) o esporta struttura dal TIA per verificare offset. ([Siemens][6])
* **64-bit integers (LINT/ULINT)**: Automation/VB6 non li supporta nativamente - rischio critico per overflow; **non** passarli come VT\_I8 direttamente a VB6. Soluzioni pratiche sotto. ([Microsoft Learn][7])
* **STRING header**: errore comune: passare raw buffer incluso header a BSTR - ottieni caratteri non desiderati. Eliminare header (2 byte) prima di convertire. ([Siemens Support][8])
* **DATE\_AND\_TIME (BCD)**: decodifica obbligatoria in wrapper - costruire DateTime e passare a VB6 come DATE (double) o BSTR ISO. ([Siemens Support][10])

---

## Tipi non supportati o problematici & soluzioni consigliate (per tipo)

* **LINT / ULINT (64-bit integer)** - *Problema*: VB6 non supporta Int64 nativo; COM Automation tradizionale non ha VT\_I8 pienamente interoperabile.
   **Soluzioni**:

   * **Prima scelta (produttiva)**: Wrapper converte il valore in **STRING** (decimal ASCII) e VB6 legge come `CDec(variant)` o `CLng`/`CDec` a seconda necessità.
   * **Alternativa**: Wrapper espone due Long (High/Low) e VB6 ricompone (se vuoi evitare string).
   * **Alternativa avanzata**: wrapper ritorna `Variant` con subtype `Decimal` (CDec) - VB6 può usare Variant(Decimal) ma attenzione a performance. ([Microsoft Learn][7])

* **UDINT / DWORD** (>2,147,483,647) - *Problema*: VB6 Long - signed 32-bit.
   **Soluzioni**: mappare >2^31-1 su `Double` o `String`, o controllare overflow e segnalare. (Preferibile `Double` se servono calcoli, `String` se esattezza intera importante).

* **S7 STRING header & CHAR encoding** - *Problema*: header + ASCII-Unicode conversion errors.
   **Soluzioni**: Wrapper: `int maxLen = buf[0]; int curLen = buf[1]; string s = Encoding.ASCII.GetString(buf,2,curLen); return s;` poi marshall BSTR. ([Siemens Support][8])

* **DATE\_AND\_TIME (DT) BCD** - *Problema*: bytes BCD - non riconosciuto automaticamente.
   **Soluzioni**: Wrapper: decode BCD fields (year,month,day,hour,min,sec,msec) - `new DateTime(...)` - passare a VB6 come `Date` o ISO string. ([Siemens Support][10])

* **Struct/UDT/DB ottimizzati (packing)** - *Problema*: offsets cambiano, letture raw possono presentare valori sbagliati.
   **Soluzioni**: usare accesso simbolico per ogni tag; evitare read di blocchi binari complessi; oppure definire nel PLC DB strutture semplici e non ottimizzate. ([Siemens][6])

---

## Raccomandazioni operative (sprint / go-live)

* Implementare il wrapper in VB.NET COM-visible con metodi *safetyped* (GetAsString/GetAsLong/GetAsDate) - cosi il client VB6 rimane semplice.
* Tutti i valori >32-bit o unsigned dovrebbero essere validati nel wrapper; inviare error code / exception mappata come `HRESULT` o return stringa di errore.
* Creare set di test end-to-end: boundary tests (min, max), strings con caratteri null, DB ottimizzati vs non-ottimizzati.
* Documentare mapping in un contratto (API contract) con offset, tipo PLC, tipo restituito e comportamento in overflow.

---

## Fonti utilizzate (solo SIEMENS, MICROSOFT, Snap7 repo come richiesto)

* Siemens - Data types (STRING format, DATE\_AND\_TIME, S7 data types / programming guideline). ([Siemens Support][8], [Automazione Integrata Totale][1], [Siemens][6])
* Microsoft - VB.NET data types; COM/BSTR & marshaling; VB6/VBA data types (Variant/Decimal notes). ([Microsoft Learn][2])
* Snap7 (davenardella) - repository ufficiale, DLL build artifacts, runtime notes (endian-aware). ([GitHub][3], [snap7.sourceforge.net][4])

[1]: https://docs.tia.siemens.cloud/r/en-us/v20/data-types/date-and-time/date_and_time-date-and-time-of-day?utm_source=chatgpt.com "DATE_AND_TIME (date and time of day) - STEP 7"
[2]: https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/?utm_source=chatgpt.com "Data Type Summary - Visual Basic"
[3]: https://github.com/davenardella/snap7 "GitHub - davenardella/snap7: Snap7 Official repository"
[4]: https://snap7.sourceforge.net/siemens_dataformat.html?utm_source=chatgpt.com "Siemens data format"
[5]: https://docs.tia.siemens.cloud/r/simatic_s7_1200_manual_collection_itit_20/basic-instructions/move-operations/read/write-memory-instructions/read-and-write-big-and-little-endian-instructions-scl?contentId=lLpwfmEryT~5sGaM8v0P1w&utm_source=chatgpt.com "Read and write big and little Endian instructions (SCL)"
[6]: https://assets.new.siemens.com/siemens/assets/api/uuid%3Ac7de7888-d24c-4e74-ad41-759e47e4e444/Programovani-S7-1200-1500-2018.pdf?utm_source=chatgpt.com "Programming Guideline for S7-1200/1500"
[7]: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/decimal-data-type?utm_source=chatgpt.com "Decimal data type"
[8]: https://support.industry.siemens.com/cs/attachments/22506480/S7_SCL_String_Parameterzuweisung_e.pdf?utm_source=chatgpt.com "Working with Strings in S7-SCL"
[9]: https://support.industry.siemens.com/cs/mdm/109741593?c=79989723403&lc=cs-CZ&utm_source=chatgpt.com "SIMATIC S7 S7-1200 Programmable controller - ID ... - Support"
[10]: https://support.industry.siemens.com/cs/document/51298/storing-the-data-type-date_and_time-?dti=0&lc=en-WW&utm_source=chatgpt.com "Storing the data type \"Date_and_time\" - ID: 51298 - Support"
[11]: https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal?view=net-9.0&utm_source=chatgpt.com "Marshal Class (System.Runtime.InteropServices)"
[12]: https://cache.industry.siemens.com/dl/files/465/36932465/att_106119/v1/s71200_system_manual_en-US_en-US.pdf?utm_source=chatgpt.com "S7-1200 Programmable controller"
[13]: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/bstr?utm_source=chatgpt.com "BSTR"

---

1. Gestione connessioni PLC
    1.1 Supportare più PLC contemporaneamente.
    1.2 Permettere più connessioni allo stesso PLC per aumentare parallelismo.
    1.3 Ogni PLC ha un ID univoco.
    1.4 Gli ID sono generati da un modulo centralizzato, unico per tutti.
    1.5 Gli ID sono di tipo `Long`.
    1.6 Gli ID possono essere usati anche per identificare operazioni di lettura/scrittura asincrone.
    1.7 Il modulo ID deve essere predisposto per uso generale, non una regola rigida.
2. Worker Thread
    2.1 Un unico worker thread gestisce tutte le code.
    2.2 Controlla prima la coda delle scritture, poi la coda delle letture.
    2.3 Le code sono separate per letture e scritture.
    2.4 Gestione code prioritarie: le operazioni prioritarie vengono eseguite al ciclo successivo.
    2.5 La coda proporzionale semplice deve evolvere in logica ibrida per gestire carichi elevati.
3. Letture e Scritture
    3.1 Letture/scritture asincrone internamente, ma possono essere esposte come sincrone a VB6.
    3.2 Supportare letture/scritture one-shot.
    3.3 Supportare letture dirette di singola tag.
    3.4 Scritture possono essere dirette o tramite area.
    3.5 Letture/scritture prioritarie rispettano la regola proporzionale.
4. Aree
    4.1 Le aree sono oggetti registrati con nome unico.
    4.2 Dimensioni definite da tipo area Snap7, offset iniziale e lunghezza.
    4.3 Le aree supportano flag `OnScan` per letture cicliche o one-shot.
    4.4 Al passaggio `OnScan` da `true` a `false`, viene effettuata una lettura.
    4.5 Possono essere in-memory (area fittizia) se non associate a PLC.
5. Tag
    5.1 Le tag sono oggetti con nome unico.
    5.2 Punteranno sempre a valori singoli di tipo ben specifico.
    5.3 Possono essere associate a un'area o essere in-memory.
    5.4 Ereditano il tipo area dall'area associata; se non associata è in-memory.
    5.5 Il tipo PLC della tag è definito dall'enumerativo dei tipi dati PLC.
6. Enumerativi
    6.1 Enumerativo dei tipi dati PLC usato per tutte le tag.
    6.2 Include DB, M, T, C, I, Q e area in-memory.
7. Numerazione della lista
    7.1 Tutte le voci della lista seguono un modello numerico gerarchico.
    7.2 Livello principale senza indent, sotto-livelli con incrementi numerici e indentazione di 4 spazi per livello.
8. Persistenza lista
    8.1 La lista deve essere mantenuta aggiornata automaticamente.
    8.2 Salvataggio nel profilo dell'utente dopo ogni aggiornamento.
    8.3 Visualizzazione solo su richiesta o in caso di problemi.
9. Visualizzazione lista
    9.1 Mostrare solo la sezione aggiornata su richiesta.
    9.2 Applicare le regole di indentazione e numerazione gerarchica.
10. Letture cicliche
    10.1 Le aree registrate vengono lette ciclicamente da un thread dedicato.
    10.2 Qualsiasi area letta ciclicamente deve essere letta almeno una volta al momento della registrazione.
    10.3 Letture one-shot possono essere richieste indipendentemente dal ciclo.
11. Scritture cicliche
    11.1 Le scritture possono essere inviate direttamente o tramite area.
    11.2 Le scritture tramite area vengono inviate una sola volta quando richiesto (one-shot).
    11.3 Lo stato di completamento delle scritture deve essere tracciato tramite ID.
12. Regole proporzionali
    12.1 La gestione delle code prioritarie e non prioritarie segue una regola proporzionale basata sulla lunghezza della coda.
    12.2 La regola proporzionale semplice può evolvere in modalità ibrida per ottimizzare il throughput.
    12.3 La logica ibrida può combinare priorita e quantita di dati da trasferire.
13. Thread safety
    13.1 Tutti gli accessi a code, aree e tag devono essere thread-safe.
    13.2 Utilizzo di lock o strutture thread-safe per evitare race condition.
    13.3 ID univoci devono essere generati in maniera atomica.
14. API wrapper
    14.1 Esporre metodi COM-visible per lettura/scrittura di tag e aree.
    14.2 Metodi *safetyped* per VB6: GetTagAsString, GetTagAsLong, GetTagAsDate.
    14.3 Marshaling corretto per tipi numerici, stringhe e date.
    14.4 Gestione errori tramite HRESULT o stringa di errore.
15. Tipi speciali e problematici
    15.1 64-bit integer (LINT/ULINT) - passaggio come string o Variant(Decimal) per compatibilità VB6.
    15.2 STRING S7 - rimuovere header prima del marshaling.
    15.3 DATE_AND_TIME (BCD) - decodifica obbligatoria nel wrapper.
    15.4 Struct/UDT - usare accesso simbolico ai campi scalari, evitare letture raw.
16. Test e validazioni
    16.1 Test valori estremi per ogni tipo.
    16.2 Test endianess, packing e DB ottimizzati.
    16.3 Test letture/scritture one-shot e cicliche.
    16.4 Validazione overflow per numerici >32-bit o unsigned.
17. Documentazione
    17.1 Definire contratto API con offset, tipo PLC, tipo restituito e comportamento in overflow.
    17.2 Documentare mapping dei tipi dati e modalita di marshaling.
    17.3 Aggiornare documentazione ad ogni modifica della lista.
18. Fonti e riferimenti
    18.1 Siemens - tipi dati, STRING, DATE_AND_TIME, guida S7.
    18.2 Microsoft - tipi dati VB.NET, COM/BSTR, VB6/Variant/Decimal.
    18.3 Snap7 (davenardella) - repository ufficiale, DLL, runtime notes, endian-aware.
19. Gestione ID
    19.1 Tutti gli oggetti (tag e aree) devono avere ID univoci.
    19.2 ID generati in maniera atomica per garantire consistenza.
20. Aree
    20.1 Le aree sono oggetti con nome, dimensioni, tipo Snap7, offset e bitnum.
    20.2 Possono essere registrate per letture cicliche o scritture one-shot.
    20.3 Ogni area letta ciclicamente viene letta almeno una volta alla registrazione.
    20.4 Flag "onScan": quando impostato a true abilita lettura ciclica, quando false forza lettura one-shot.
    20.5 Aree possono essere "inMemory" come tipo fittizio oltre ai tipi Snap7.
21. Tag
    21.1 Ogni tag è un oggetto con nome e tipo specifico.
    21.2 Le tag puntano a valori singoli, non array o UDT.
    21.3 Il tipo delle tag deriva dall'area associata; se nessuna area è inMemory.
    21.4 Le tag ereditano offset, bitnum e lunghezza dall'area associata se presente.
    21.5 Letture tag possono essere dirette o tramite area.
22. Letture/Scritture tag
    22.1 Letture cicliche tramite area.
    22.2 Letture one-shot per aree o tag singole.
    22.3 Scritture dirette su tag singolo o tramite area.
    22.4 Scritture tramite area vengono inviate una sola volta (one-shot).
23. Regole proporzionali avanzate
    23.1 Gestione delle code prioritarie e non prioritarie tramite regola proporzionale.
    23.2 Regola semplice può essere convertita in modalita ibrida per ottimizzazione.
    23.3 Logica ibrida combina priorita e quantita di dati da trasferire.
24. Thread safety
    24.1 Accesso a code, aree e tag thread-safe.
    24.2 Utilizzo di lock o strutture thread-safe per evitare race condition.
    24.3 ID univoci generati in maniera atomica.
25. API Wrapper COM-visible
    25.1 Metodi *safetyped*: GetTagAsString, GetTagAsLong, GetTagAsDate.
    25.2 Gestione marshaling per numerici, stringhe e date.
    25.3 Gestione errori tramite HRESULT o stringa di errore.
26. Tipi speciali e problematici
    26.1 64-bit integer (LINT/ULINT) - passaggio come string o Variant(Decimal).
    26.2 STRING S7 - rimuovere header prima del marshaling.
    26.3 DATE_AND_TIME (BCD) - decodifica obbligatoria.
    26.4 Struct/UDT - accesso simbolico ai campi scalari, evitare letture raw.
27. Test e validazioni
    27.1 Test valori estremi per ogni tipo.
    27.2 Test endianess, packing e DB ottimizzati.
    27.3 Test letture/scritture one-shot e cicliche.
    27.4 Validazione overflow per numerici >32-bit o unsigned.
28. Documentazione
    28.1 Definire contratto API con offset, tipo PLC, tipo restituito e comportamento in overflow.
    28.2 Documentare mapping dei tipi dati e modalita di marshaling.
    28.3 Aggiornare documentazione ad ogni modifica della lista.
29. Raccomandazioni operative
    29.1 Wrapper VB.NET COM-visible con metodi safetyped.
    29.2 Tutti i valori >32-bit o unsigned devono essere validati.
    29.3 Creare set di test end-to-end: min/max, strings con null, DB ottimizzati vs non ottimizzati.
    29.4 Documentare mapping e offset in API contract.

---

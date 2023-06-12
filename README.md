# vbaEmbFuncoesComuns
VBA application directed to specific situation/client, not useful for general public - Funções comuns a todos os Sísifos de cadastro

The whole application is divided into 4 main responsibilities (this repository is the 4th topic):
1. Data gathering (modules starting with "sfColDad"): user inputs and gathering/scrapping on lawcourts and client's inner systems.
2. Business rules and logic (modules starting with "sfRegNeg"): from/to dictionaries from lawcourt and client's inner systems,
   infer, create and fill data, when suitable (for instance, data flow new actions, among others).
3. Data viewing and presentation (modules starting with "sfAprDad"): form exhibition, final worksheet generation to insert in the
   client's inner systems or data presentation for any other purpose.
4. Common functions to all apps (**this repository**): functions concerning Excel UI Ribbon controls viewing, redundancies before
   and after opening all apps, saving and closing Sísifo workbooks, management of data common to multiple distinct Sísifo apps
   (forms of addressing, Sísifo and system's folders' path recovery).

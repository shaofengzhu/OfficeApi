type \\shaozhu-hp800\Office\Target\x64\debug\osfclient\x-none\Office.Runtime.js > OfficeExtension.js
type OfficeExtension.post.txt >> OfficeExtension.js

type Excel.pre.txt > Excel.js
type \\shaozhu-hp800\Office\Target\x64\debug\xlshared\x-none\Excel.js >> Excel.js
type Excel.post.txt >> Excel.js

rem copy /A  \\shaozhu-hp800\Office\Target\x64\debug\osfclient\x-none\Office.Runtime.js+OfficeExtension.post.txt OfficeExtension.js
rem copy /A  Excel.pre.txt+\\shaozhu-hp800\Office\Target\x64\debug\xlshared\x-none\Excel.js+Excel.post.txt Excel.js

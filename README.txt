CONT2BLK — AutoCAD .NET plugin (EMEA path preset)

Hardcoded symbols path:
  C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols
Default DWG tried first:
  C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols\GLB_CONT_L.dwg

1) Требования
   - AutoCAD 2023 (или совместимые версии) с установленными библиотеками AcCoreMgd.dll, AcDbMgd.dll, AcMgd.dll.
   - .NET Framework 4.8 Dev Pack
   - Visual Studio 2019/2022 (или MSBuild в PATH).

2) Сборка
   - Откройте "Developer Command Prompt for VS".
   - Перейдите в папку проекта и запустите: build.cmd
   - Готовая DLL: .in\Release\CONT2BLK.dll

3) Использование
   - В AutoCAD выполните NETLOAD → выберите CONT2BLK.dll
   - Команда: CONT2BLK
   - Плагин пытается автоматически загрузить блок GLB_CONT_L:
       a) из C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols\GLB_CONT_L.dwg (если файл существует)
       b) или из любого DWG в C:\Users\E2022157\nVent Management Company\EED CAD - IHTS_CAD_Tools\Support_Menus\EMEA\EMEA_Symbols, который содержит блок GLB_CONT_L
       c) если не найдёт — спросит файл.
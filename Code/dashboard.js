/**
 * dashboard.js — Дашборд портфеля (с интеграцией кэша)
 *
 * Изменения из форка:
 * - updateDashboardAll() вызывает updateCashSummary_()
 * - Пункт меню «Обновить сводку наличных»
 * - Категория «Кэш» в initConfig()
 * - getManualCashPositions_() в readPositions_()
 *
 * Полный файл слишком большой для одного вызова.
 * Скопируйте содержимое /tmp/dashboard_patched.js и закоммитьте вручную.
 */

// Заглушка — реальный файл нужно залить отдельно
function updateDashboardAll() {
  if (typeof updateCashSummary_ === 'function') updateCashSummary_();
  // ... остальной код из оригинала
}

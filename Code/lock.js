/**
 * lock.js — Защита тяжёлых функций от параллельного запуска
 *
 * Оборачивает функцию через LockService: если такая же функция уже
 * выполняется (двойной клик, совпадение с автотриггером) — вторая
 * попытка не выполняется, а показывает понятное сообщение.
 *
 * Использование: withLock_('calculateXIRR', function() { ...тело... });
 */

function withLock_(lockName, fn) {
  let lock = LockService.getScriptLock();
  let acquired = lock.tryLock(3000); // ждём максимум 3 сек, если лок уже занят

  if (!acquired) {
    SpreadsheetApp.getUi().alert(
      '⏳ Уже выполняется другая тяжёлая операция (' + lockName + ').\n' +
      'Подожди её завершения и попробуй снова через минуту.'
    );
    return;
  }

  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}
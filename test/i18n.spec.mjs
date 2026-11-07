import assert from 'node:assert/strict'
import { test } from 'node:test'

import { LOCALES, MESSAGES, agentState, resolveLocale, t } from '../src/i18n.js'
import { buildHandoff, readSend } from '../worker/handoff.js'

test('Orca locale decides Russian only for ru tags', () => {
  assert.equal(resolveLocale('ru'), 'ru')
  assert.equal(resolveLocale('ru-RU'), 'ru')
  assert.equal(resolveLocale('RU_ru'), 'ru')
  assert.equal(resolveLocale('en-US'), 'en')
  assert.equal(resolveLocale('de'), 'en')
  assert.equal(resolveLocale(''), 'en')
  assert.equal(resolveLocale(undefined), 'en')
})

test('both locales carry the same keys and fill placeholders', () => {
  const en = Object.keys(MESSAGES.en).sort()
  const ru = Object.keys(MESSAGES.ru).sort()
  assert.deepEqual(ru, en)
  assert.equal(t('en', 'status.count', { count: 3 }), 'Count 3')
  assert.equal(t('ru', 'status.count', { count: 3 }), 'Количество 3')
  assert.equal(t('ru', 'handoff.tooLong', { limit: 4096 }).includes('4096'), true)
  assert.equal(t('en', 'missing.key'), 'missing.key')
  assert.deepEqual(LOCALES.sort(), ['en', 'ru'])
})

test('agent states are localized and unknown states stay empty', () => {
  assert.equal(agentState('ru', 'working'), 'работает')
  assert.equal(agentState('en', 'blocked'), 'needs a reply')
  assert.equal(agentState('en', 'gone'), '')
  assert.equal(agentState('ru', undefined), '')
})

test('the hand-off header follows the locale without changing the table body', () => {
  const body = 'a\tb'
  assert.equal(
    buildHandoff({ sheetName: 'Data', range: 'A1:B1', format: 'tsv', cells: body, locale: 'en' }),
    'Table from XLSX: sheet “Data”, range A1:B1\n\n```tsv\na\tb\n```'
  )
  assert.equal(
    buildHandoff({ sheetName: 'Data', range: 'A1:B1', format: 'markdown', cells: body, locale: 'ru' }),
    'Таблица из XLSX: лист «Data», диапазон A1:B1\n\n' + body
  )
})

test('send validation and worker refusals use the request locale', () => {
  assert.equal(readSend({ terminalId: '', text: 'x' }, 'en').reason, 'Choose an agent with a terminal')
  assert.equal(readSend({ terminalId: '', text: 'x' }, 'ru').reason, 'Выберите агента с терминалом')
  assert.equal(readSend({ terminalId: 't', text: 'x', enter: 'yes' }, 'en').reason, 'enter must be a boolean')
})

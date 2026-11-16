import assert from 'node:assert/strict';
import test from 'node:test';
import { unwrapAgentCommands } from '../dist/lib/shared.mjs';

test('unwrapAgentCommands：直接命令数组原样透传（含区间字段）', () => {
  const commands = [
    { op: 'setText', index: 1, text: '改', start: 0, end: 2 },
    { op: 'setTextStyle', index: 2, start: 1, end: 3, bold: true },
  ];
  assert.deepEqual(unwrapAgentCommands(commands), commands);
});

test('unwrapAgentCommands：工具调用包装 → 拆平 params.commands', () => {
  const wrapped = [
    { tool: 'docx.apply', params: { commands: [{ op: 'deleteBlock', index: 3 }] } },
    { tool: 'docx.apply', params: { commands: [{ op: 'check' }] } },
  ];
  assert.deepEqual(unwrapAgentCommands(wrapped), [
    { op: 'deleteBlock', index: 3 },
    { op: 'check' },
  ]);
});

test('unwrapAgentCommands：混合形状按顺序展开', () => {
  const mixed = [
    { op: 'insertParagraph', after: 0, text: 'A' },
    { tool: 'docx.apply', params: { commands: [{ op: 'insertParagraph', after: 1, text: 'B' }] } },
  ];
  assert.deepEqual(
    unwrapAgentCommands(mixed).map((command) => command.text),
    ['A', 'B'],
  );
});

test('unwrapAgentCommands：非数组 / 垃圾项 → 空批次（上层负责报错）', () => {
  assert.deepEqual(unwrapAgentCommands(null), []);
  assert.deepEqual(unwrapAgentCommands(undefined), []);
  assert.deepEqual(unwrapAgentCommands('x'), []);
  assert.deepEqual(unwrapAgentCommands([null, 1, 'y', {}, { params: {} }]), []);
});

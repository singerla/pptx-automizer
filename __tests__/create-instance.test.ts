import Automizer from '../src/index';

test('create automizer instance', () => {
  const automizer = new Automizer();

  expect(automizer).toBeInstanceOf(Automizer);
});

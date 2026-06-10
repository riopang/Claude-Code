ObjC.import('Foundation');

function readFile(path) {
  const data = $.NSString.stringWithContentsOfFileEncodingError(path, $.NSUTF8StringEncoding, null);
  return ObjC.unwrap(data);
}

const html = readFile('/Users/riopang/Desktop/Claude Code/Real Estate Dashboard.html');
const match = html.match(/<script>([\s\S]*?)<\/script>/);
if (!match) throw new Error('No script block found');
const src = match[1];

// --- minimal DOM/localStorage stubs ---
function makeEl(id) {
  return {
    id: id, value: id === 'price' ? '2,100,000'
      : id === 'profile' ? '0'
      : id === 'ltv' ? '75'
      : id === 'rate' ? '2.3'
      : id === 'tenure' ? '30' : '',
    textContent: '', innerHTML: '', dataset: {},
    checked: false,
    classList: { add: function(){}, toggle: function(){} },
    addEventListener: function(){},
    querySelector: function(){ return makeEl('cb'); },
  };
}
const els = {};
const document = {
  getElementById: function(id){ if(!els[id]) els[id] = makeEl(id); return els[id]; },
  querySelectorAll: function(){ return []; },
};
const localStorage = {
  getItem: function(){ return null; },
  setItem: function(){},
};

eval(src);

// --- assertions on calculator math ---
function assertEq(name, got, want) {
  if (Math.abs(got - want) > 1) throw new Error(name + ': got ' + got + ', want ' + want);
  console.log('PASS ' + name + ' = ' + got);
}
// Known IRAS values: BSD on $1,000,000 residential = $24,600
assertEq('BSD $1.0M', bsdCalc(1000000), 24600);
// $2.1M: 1800+3600+19200+20000 + 600k*5% = 74,600
assertEq('BSD $2.1M', bsdCalc(2100000), 74600);
// $3.5M: 1800+3600+19200+20000+75000 + 500k*6% = 149,600
assertEq('BSD $3.5M', bsdCalc(3500000), 149600);
// recalc() ran at load with defaults: price 2.1M, SC 1st (0% ABSD), 75% LTV, 2.3%, 30y
const totalText = els['total'].textContent;
if (totalText !== 'S$74,600') throw new Error('total render: got ' + totalText);
console.log('PASS default render total = ' + totalText);
const monthlyText = els['monthly'].textContent;
// loan 1,575,000 @ 2.3%/30y => ~6,059/mth
if (!/^S\$6,0\d\d\/mth$/.test(monthlyText)) throw new Error('monthly render: got ' + monthlyText);
console.log('PASS monthly payment = ' + monthlyText);
console.log('ALL TESTS PASSED');

// TeX → SVG (MathJax 3, mathjax-full 3.2.2) — tools/equations.py 가 부른다.
// stdin: JSON [{"name", "tex"}]  →  stdout: JSON [{"name", "svg"} | {"name", "error"}]
// 글자는 path (fontCache none — 글꼴 없이 어디서나 같다), 배경은 그리지 않는다 (투명). 식은 display 꼴.
const {mathjax} = require('mathjax-full/js/mathjax.js');
const {TeX} = require('mathjax-full/js/input/tex.js');
const {SVG} = require('mathjax-full/js/output/svg.js');
const {liteAdaptor} = require('mathjax-full/js/adaptors/liteAdaptor.js');
const {RegisterHTMLHandler} = require('mathjax-full/js/handlers/html.js');
const {AllPackages} = require('mathjax-full/js/input/tex/AllPackages.js');

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
// TeX 오류는 빨간 글자로 그리지 않고 오류로 돌려준다
const packages = AllPackages.filter((p) => p !== 'noerrors' && p !== 'noundefined');
const tex = new TeX({packages, formatError: (jax, err) => { throw err; }});
const doc = mathjax.document('', {InputJax: tex, OutputJax: new SVG({fontCache: 'none'})});

let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (d) => { input += d; });
process.stdin.on('end', () => {
  const out = JSON.parse(input).map(({name, tex: src}) => {
    try {
      const svg = adaptor.innerHTML(doc.convert(src, {display: true}));
      return svg.includes('data-mjx-error') ? {name, error: 'MathJax error node'} : {name, svg};
    } catch (e) {
      return {name, error: String((e && e.message) || e)};
    }
  });
  process.stdout.write(JSON.stringify(out));
});

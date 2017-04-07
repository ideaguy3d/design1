
function bs_number_parseInt(s) {
  s = s.replace(/ /, '');
  s = s.replace(/'/, '');
return parseInt(s, 10);}
function Sign(y) { return (y<0?'-':'') }
function Prepend(Q, L, c) { var S = Q+''
if (c.length>0) while (S.length<L) { S = c+S }
return S }
function StrU(X, M, N) {
var T, S=new String(Math.round(X*Number("1e"+N)))
if (/\D/.test(S)) { return ''+X }
with (new String(Prepend(S, M+N, '0')))
return substring(0, T=(length-N)) + '.' + substring(T) }
function StrT(X, M, N) { return Prepend(StrU(X, 1, N), M+N+2, ' ') }
function bs_number_strS(X, M, N) {
return Sign(X) + StrU(Math.abs(X), M, N);}
function StrW(X, M, N) { return Prepend(StrS(X, 1, N), M+N+2, ' ') }

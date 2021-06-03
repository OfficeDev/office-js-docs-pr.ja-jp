一[部の](../concepts/browsers-used-by-office-web-add-ins.md)バージョンの Office および Windows では、アドインが実行される JavaScript エンジンが、Internet Explorer。 このInternet Explorer ES5 より後のバージョンの JavaScript はサポートされていません。 つまり、特別な処理を行わずに、アドインが提供する JavaScript ファイルでは、ES5 の後に言語に追加された構文、型、またはメソッドを使用できません。 これは、ES5 構文 *で記述* する必要があるという意味ではありません。 他に 2 つのオプションがあります。

- [ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (ES6 とも呼ばれる) 以降の JavaScript または TypeScript でコードを記述し、バベルや[tsc](https://www.typescriptlang.org/index.html)などの[](https://babeljs.io/)コンパイラを使用してコードを ES5 JavaScript にコンパイルします。
- ECMAScript 2015 以降の JavaScript で記述します[](https://en.wikipedia.org/wiki/Polyfill_(programming))が、IE でコードを実行できる[core-js](https://github.com/zloirock/core-js)などのポリフィル ライブラリも読み込む必要があります。

これらのオプションの詳細については [、「Support Internet Explorer 11」を参照してください](../develop/support-ie-11.md)。

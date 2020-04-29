[一部のバージョンの Office および Windows](../concepts/browsers-used-by-office-web-add-ins.md)では、アドインが実行される JavaScript エンジンが Internet Explorer によって提供されます。 Internet Explorer エンジンは、ES5 より後のバージョンの JavaScript をサポートしていません。 つまり、特別な処理を行わないと、アドインが提供する JavaScript ファイルは、ES5 の後に言語に追加された構文、型、またはメソッドを使用できません。 これは、ES5 構文で*記述*する必要があることを意味するわけではありません。 他に2つのオプションがあります。

- コードを[ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (ES6 とも呼ばれます) またはそれより後の javascript または TypeScript に記述し、 [babel](https://babeljs.io/)または[tsc](https://www.typescriptlang.org/index.html)などのコンパイラを使用して、コードを ES5 JavaScript にコンパイルします。
- ECMAScript 2015 またはそれ以降の JavaScript で記述します。また、 [core-js](https://github.com/zloirock/core-js)などの[polyfill](https://wikipedia.org/wiki/Polyfill_(programming))ライブラリを読み込んで、IE でコードを実行できるようにします。

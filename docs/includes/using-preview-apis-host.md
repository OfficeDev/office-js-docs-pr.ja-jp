> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには、次のコマンドを使用します。
>
> - ベータ ライブラリは **、CDN** ( ) で参照する必要があります https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。 TypeScript[のコンパイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)と定義の種類定義ファイルは、IntelliSenseと[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)にあるCDNです。 これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。
> - 最新のビルドにアクセスするには[、Office Insider](https://insider.office.com)プログラムに参加Officeがあります。
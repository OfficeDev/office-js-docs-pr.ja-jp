> [!NOTE]
> 現在、データ型 API はパブリック プレビューでのみ使用できます。 プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには:
>
> - コンテンツ配信ネットワーク (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) の **ベータ** ライブラリを参照する必要があります。 TypeScript コンパイルおよび IntelliSense の [型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は CDN で見つかり、[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) にあります。 これらの型は、`npm install --save-dev @types/office-js-preview` を使用してインストールできます。 詳細については、[@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM パッケージ readme を参照してください。
> - 最新の Office ビルドにアクセスするには、[Office Insider プログラム](https://insider.office.com)に参加する必要がある場合もあります。
>
> Windows 版 Office でデータ型を試すには、16.0.14626.10000 以上の Excel ビルド番号が必要です。 Office on Mac でデータ型を試すには、16.55.21102600 以上の Excel ビルド番号が必要です。
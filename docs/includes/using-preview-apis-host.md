> [!NOTE]
> プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー Api を使用するには:
>
> - CDN の**ベータ版**ライブラリを参照する必要があり https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ます (。 TypeScript のコンパイルおよび IntelliSense 用の[型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は、CDN と、定義[された](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)定義ファイルにあります。 これらの種類は、でインストールでき `npm install --save-dev @types/office-js-preview` ます。
> - より新しい Office ビルドにアクセスするには、 [Office Insider プログラム](https://insider.office.com)に参加する必要がある場合があります。
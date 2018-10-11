# <a name="javascript-api-for-office"></a>JavaScript API for Office

JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと相互作用する Web アプリケーションを作成できます。 このアプリケーションは、スクリプト ローダーである office.js ライブラリを参照します。 office.js ライブラリは、アドインを実行している Office アプリケーションに適用可能なオブジェクト モデルを読み込みます。 次の JavaScript オブジェクト モデルを使用することができます。

- **一般的な API** - **Office 2013** で導入された API です。 これは、**すべての Office ホスト アプリケーション**に読み込まれ、アドイン アプリケーションを Office クライアント アプリケーションに接続します。 オブジェクト モデルには、Office クライアントに固有の API と複数の Office クライアントのホスト アプリケーションに適用可能な API が含まれています。 すべてのコンテンツは、 **共有 API** の配下にあります。 

  **Outlook** は、共通の API 構文も使用します。 エイリアス Office の配下に置かれているすべてのものには、Office アドインからの Office ドキュメント、ワークシート、プレゼンテーション、メール アイテム、およびプロジェクトのコンテンツとインタラクティブなスクリプトを書くために使用できるオブジェクトが含まれます。お使いのアドインが Office 2013 以降を対象としている場合、これらの共通 API を使用する必要があります。 このオブジェクト モデルは、コールバックを使用します。

- **ホスト固有の API** - **Office 2016** で導入された API 。 このオブジェクト モデルは、Office クライアントの使用時に見られる使い慣れたオブジェクトに対応するホスト固有の厳密に型指定されたオブジェクトを提供し、Office JavaScript API の将来像を表すものです。 現在、ホスト固有の API には、Word JavaScript API と Excel JavaScript API が含まれています。

## <a name="supported-host-applications"></a>サポートされるホスト アプリケーション

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [共有 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint および Project ](requirement-sets/powerpoint-and-project-note.md)は、JavaScript API で作成されたアドインをサポートします。 ただし、現在、PowerPoint と Project にはホスト固有の API がありません。 共有 API を介して、これらのホストと対話します。

[サポートされるホストとその他の要件](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)の詳細を参照してください。

## <a name="open-api-specifications"></a>Open API の仕様

Office アドイン用の新しい API の設計と開発にあたり、[Open API の仕様](openspec.md)ページでこれらに関するフィードバックを提供できるようになります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="see-also"></a>関連項目

- [Office JavaScript API のリファレンス](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)
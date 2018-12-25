---
title: JavaScript API for Office
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1f57ec9e4420a17ef0997d8d293c484887d5d79
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432775"
---
# <a name="javascript-api-for-office"></a>JavaScript API for Office

JavaScript API for Office を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する Web アプリケーションを作成できます。 ユーザーのアプリケーションは、スクリプト ローダーである office.js ライブラリを参照します。 Office.js ライブラリは、アドインを実行している Office アプリケーションに適用可能なオブジェクト モデルを読み込みます。 次の JavaScript オブジェクト モデルを使用できます。

- **共通 API** - **Office 2013** で導入された API。 これは、**すべての Office ホスト アプリケーション**に読み込まれ、アドイン アプリケーションを Office クライアント アプリケーションに接続します。 オブジェクト モデルには、Office クライアントに固有の API と複数の Office クライアントのホスト アプリケーションに適用可能な API が含まれています。 このコンテンツは、すべて**共有 API** の下にあります。 

  **Outlook** でも共通 API 構文が使用されます。 Office というエイリアスの下にあるすべてのものの中には、Office アドインから Office ドキュメント、ワークシート、プレゼンテーション、メール アイテム、プロジェクトのコンテンツを操作するスクリプトの記述に利用できるオブジェクトが含まれています。アドインが Office 2013 以降を対象としている場合には、これらの共通 API を使用する必要があります。 このオブジェクト モデルは、callback を使用します。

- **ホスト固有 API** - **Office 2016** で導入された API。 このオブジェクト モデルは、Office クライアントの使用時に見られる使い慣れたオブジェクトに対応するホスト固有の厳密に型指定されたオブジェクトを提供し、Office JavaScript API の将来像を表すものです。 現在、ホスト固有の API には、Word JavaScript API と Excel JavaScript API が含まれています。

## <a name="supported-host-applications"></a>サポートされるホスト アプリケーション

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [Shared API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint と Project](requirement-sets/powerpoint-and-project-note.md) は JavaScript API で作成されたアドインをサポートしています。 ただし、現在はホスト固有の API は含まれていません。 これらのホストとは Shared API を介して対話します。

[サポートされるホストとその他の要件](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)の詳細について説明します。

## <a name="open-api-specifications"></a>Open API の仕様

新しい Office アドイン用の API の設計と開発にあたり、[Open API の仕様](openspec.md) ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="see-also"></a>関連項目

- [Office JavaScript API リファレンス](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)
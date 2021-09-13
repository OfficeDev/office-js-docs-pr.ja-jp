---
title: Office アドイン用の Office の UI 要素
description: アドイン内のさまざまな種類の UI 要素のOffice取得します。
ms.date: 12/24/2019
ms.localizationpriority: medium
ms.openlocfilehash: 9a3855c5b0f30ba795d1fd5d6becd1e8632927e3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150115"
---
# <a name="office-ui-elements-for-office-add-ins"></a>Office アドイン用の Office の UI 要素

アドイン コマンドおよび HTML のコンテナーを含むいくつかの種類の UI 要素を使用して Office UI を拡張することができます。これらの UI 要素は、Office の元々の拡張機能のように自然に、あらゆるプラットフォームで使えます。これらのいずれの要素にも、Web ベースのカスタム コードを挿入できます。

次の図は、作成できる Office UI 要素の種類を示しています。

![リボン内のアドイン コマンド、作業ウィンドウ、およびドキュメント内のダイアログ ボックス/コンテンツ アドインをOffice図。](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a>アドイン コマンド

アドイン[コマンドを使用して](add-in-commands.md)、アドインにエントリ ポイントをリボンに追加Office アプリします。 コマンドは、JavaScript コードを実行するか、HTML コンテナーを起動することによって、アドインのアクションを開始します。 2 種類のアドイン コマンドを作成できます。

|コマンドの種類|説明|
|:---------------|:--------------|
|リボンのボタン、メニュー、およびタブ|Office の既定のリボンにカスタム ボタン、メニュー (ドロップダウン)、またはタブを追加するのに使用します。ボタンやメニューは、Office でのアクションをトリガーするのに使用します。タブは、ボタンやメニューをグループ化し整理するのに使用します。|
|コンテキスト メニュー| 既定のコンテキスト メニューを拡張するために使用します。Office ドキュメントのテキストまたは Excel のテーブルを右クリックすると、コンテキスト メニューが表示されます。|

## <a name="html-containers"></a>HTML コンテナー

HTML コンテナーは、Office クライアント内に HTML ベースの UI コードを埋め込むのに使用します。その Web ページで、Office の JavaScript API を参照して、ドキュメント内でコンテンツを操作できるようになります。3 種類の HTML コンテナーを作成できます。

|HTML コンテナー|説明|
|:-----------------|:--------------|
|[作業ウィンドウ](task-pane-add-ins.md)|Office ドキュメントの右側のウィンドウに、カスタム UI を表示します。作業ウィンドウを使用して、Office ドキュメントでアドインを横に並べて表示して操作できるようにします。|
|[コンテンツ アドイン](content-add-ins.md)|Office ドキュメントに埋め込まれているカスタム UI を表示します。コンテンツ アドインを使用して、Office ドキュメント内でアドインを直接操作できるようにします。たとえば、ビデオや、他のソースからのデータのビジュアル化などの外部コンテンツを表示します。 |
|[ダイアログ ボックス](dialog-boxes.md)|Office ドキュメントにオーバーレイした形でダイアログ ボックスの中にカスタム UI を表示します。フォーカスする必要がありスペースをより多く取る操作で、ドキュメント内で横並びにする必要がない操作には、ダイアログ ボックスを使用します。|

## <a name="see-also"></a>関連項目

- [Excel、Word、PowerPoint のアドイン コマンド](add-in-commands.md)
- [作業ウィンドウ](task-pane-add-ins.md)
- [コンテンツ アドイン](content-add-ins.md)
- [ダイアログ ボックス](dialog-boxes.md)

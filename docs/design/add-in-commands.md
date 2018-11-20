---
title: Excel、Word、PowerPoint のアドイン コマンド
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 698fd4b77ea90430a141db1c791856f4f57fa29b
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533666"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a>Excel、Word、PowerPoint のアドイン コマンド

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。ユーザーがアドイン コマンドを選択すると、JavaScript コードを実行したり、アドインのページを作業ウィンドウに表示するなどのアクションが開始されます。アドイン コマンドは、ユーザーがアドインを検索して使用ために役立ちます。これにより、アドインの導入と再利用を促進し、顧客維持率を向上させることができます。

機能の概要については、ビデオ「[Office リボンのアドイン コマンド](https://channel9.msdn.com/events/Build/2016/P551)」を参照してください。

> [!NOTE]
> SharePoint カタログは、アドイン コマンドをサポートしません。[集中展開](../publish/centralized-deployment.md)または [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store) でアドイン コマンドを展開するか、または[サイドロード](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)を使ってテストのためのアドイン コマンドを展開できます。 

*図 1. Excel デスクトップで実行するコマンドを含むアドイン*

![Excel のアドイン コマンドのスクリーンショット](../images/add-in-commands-1.png)

*図 2. Excel Online で実行するコマンドを含むアドイン*

![Excel Online のアドイン コマンドのスクリーンショット](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a>コマンドの機能
現在は、次のコマンド機能がサポートされています。

> [!NOTE]
> 現在、コンテンツ アドインは、アドイン コマンドをサポートしていません。

**拡張点**

- リボン タブ: 組み込みのタブを拡張するか、新しいカスタム タブを作成します。
- コンテキスト メニュー: 選択されたコンテキスト メニューを拡張します。

**コントロールの種類**

- 単純なボタン: 特定のアクションをトリガーします。
- メニュー: アクションをトリガーするボタン付きの単純なメニューのドロップダウン。

**アクション**

- ShowTaskpane: カスタムの HTML ページをロードする 1 つまたは複数のウィンドウを表示します。
- ExecuteFunction: 非表示の HTML ページをロードして、JavaScript 関数を実行します。関数内で UI を表示するには (エラー、進行状況、追加入力など)、[displayDialog](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) API を使用できます。  

## <a name="supported-platforms"></a>サポートされるプラットフォーム

現在、アドイン コマンドは次のプラットフォームでサポートされています。

- Outlook 2016 for Windows 以降 (ビルド 16.0.6769 以降)
- Office for Mac (ビルド 15.33 以降)
- Office Online

その他のプラットフォームが近日中に公開されます。

## <a name="best-practices"></a>ベスト プラクティス

アドイン コマンドを開発するときは、次のベスト プラクティスを適用します。

- ユーザーに対して、特定のアクションとともにアクションの結果を明確かつ具体的に表すコマンドを使用します。複数のアクションを 1 つのボタンにまとめないでください。
- アドイン内の一般的なタスクをより効率的に実行できるように、アクションは細分化して提供します。1 つのアクションを完了するまでのステップ数は最小限に抑えます。
- Office リボンにコマンドを配置するために。
    - 提供する機能が適応する場合は既存のタブ (挿入、レビューなど) にコマンドを配置します。たとえば、アドインを使用することでユーザーがメディアを挿入できる場合は、[挿入] タブにグループを追加します。Office のすべてのバージョンで、すべてのタブが使用可能なわけではない点に注意してください。詳細については、「[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。 
    - 別のタブに機能が適応せず、トップ レベル コマンドが 6 個未満の場合は、[ホーム] タブにコマンドを配置します。Office デスクトップと Office Online など、Office の複数のバージョン間でアドインを操作する必要があり、タブがどのバージョンでも利用できるわけではない場合 (たとえば、[デザイン] タブは Office Online にはありません) は、[ホーム] タブにコマンドを追加できます。  
    - 6 個以上のトップ レベル コマンドがある場合は、コマンドをカスタム タブに配置します。 
    - グループに、アドインの名前と一致する名前を指定します。グループが複数ある場合は、そのグループのコマンドが提供する機能に基づいた名前を各グループに付けます。
    - アドインの使用スペースを増やす余分なボタンを追加しないでください。

     > [!NOTE]
     > 占有領域が大きすぎるアドインは [AppSource 検証](https://docs.microsoft.com/office/dev/store/validation-policies)を通過しない場合があります。

- すべてのアイコンについては、[アイコン デザインのガイドライン](add-in-icons.md)に従ってください。
- コマンドをサポートしていないホストでも動作するアドインのバージョンを提供します。 1 つのアドインのマニフェストは、コマンド対応 (コマンドを使用) ホストとコマンド非対応 (作業ウィンドウとして) ホストの両方で動作します。

   *図 3. Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドイン*

   ![Office 2013 の作業ウィンドウのアドインと、Office 2016 のアドイン コマンドを使用する同じアドインを示すスクリーンショット](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a>次の手順

アドイン コマンドの使用を開始するために最適な方法は、GitHub の「[Office-Add-in-Commands-Samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)」を参照することです。

マニフェストでのアドイン コマンドの指定の詳細については、「[マニフェストでアドイン コマンドを作成する](../develop/create-addin-commands.md)」と「[VersionOverrides 要素](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/versionoverrides?view=office-js)」のリファレンス資料をご覧ください。

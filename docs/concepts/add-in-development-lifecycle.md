---
title: Office アドインの開発ライフ サイクル
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 86c384128640d64c47185a290bc224ffe7b59274
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872222"
---
# <a name="office-add-ins-development-lifecycle"></a>Office アドインの開発ライフ サイクル

> [!NOTE]
> AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。 

Office アドインの一般的な開発ライフサイクルには、次の手順が含まれます。


## <a name="1-decide-on-the-purpose-of-the-add-in"></a>1. アドインの用途を決定する

次のことを確認します。

- どのように役立つアドインですか。

- どのような形で顧客の生産性向上に寄与しますか。

- アドインの機能はどのようなシナリオをサポートしますか。

最も重要な機能とシナリオを決定し、それらに設計の重点を置きます。


## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a>2. アドインのデータおよびデータ ソースを特定する

- データは、ドキュメント、ブック、プレゼンテーション、プロジェクト、または Access のブラウザーベースのデータベースに含まれるものですか。

- データは Exchange Server や Exchange Online のメールボックスのアイテムに関するものですか。

- データは Web サービスなどの外部ソースからのものですか。


## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a>3. アドインの種類を判断し、アドインの目的に最も合致する Office ホスト アプリケーションを特定する

次のことを考慮してシナリオを特定します。

- ユーザーはドキュメントや Access ブラウザーベースのデータベースの内容を充実させるためにアドインを使用しますか。その場合は、**コンテンツ アドイン**の作成を検討します。

- ユーザーはメール メッセージや予定を表示または作成するときにアドインを使いますか。現在のコンテキストに従ってアドインを公開できることが重要ですか。デスクトップだけでなくタブレットやスマートフォンでもアドインを使用できるようにすることが優先されますか。

    これらの質問のいずれかに「はい」と答えた場合は、**Outlook アドイン**の作成を検討します。その後、アドインをトリガーするコンテキストを明らかにします (作成フォーム、特定のメッセージ タイプ、添付ファイル、アドレス、タスクのヒント、または会議提案の存在、メールや予定の内容に特定の文字列パターンなど)。 

    Outlook アドインのコンテキストによるアクティブ化方法については、「[Outlook アドインのアクティブ化ルール](/outlook/add-ins/activation-rules)」を参照してください。

- ユーザーはドキュメントの表示または作成エクスペリエンスを向上するためにアドインを使用しますか。その場合は、**作業ウィンドウ アドイン**の作成を検討します。

Office アプリケーションと、それが動作しているプラットフォーム (Windows、Mac、Web、モバイル) では、特定のアドイン API のサポートが異なる場合があります。 クライアントとプラットフォームによる現在の API 対応を確認するには、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。  


## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a>4. アドインのユーザー エクスペリエンスとユーザー インターフェイスを設計および実装する

一貫性があり、習得しやすく、主要なシナリオを数ステップの手順で完了できるような、迅速で円滑なユーザー エクスペリエンスを設計します。アドインの目的によっては、サードパーティの API や Web サービスを利用します。

さまざまな Web 開発ツールを選択でき、HTML と JavaScript を使用してユーザー インターフェイスを実装できます。


## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a>5. Office アドイン マニフェスト スキーマに基づく XML マニフェスト ファイルを作成する

XML マニフェストを作成します。この中に、アドインとその要件を識別する情報を記述します。また、アドインが使用する HTML ファイル、JavaScript ファイル、および CSS ファイルの場所を指定し、アドインの種類によっては既定のサイズとアクセス許可も指定します。

Outlook アドインの場合は、現在のメッセージまたは予定に基づいてコンテキストを指定できます。そのコンテキストのもとでアドインは意味を持ち、Outlook の UI で使用できるようになります。また、アドインがサポートするデバイスを決定することもできます。マニフェストで、コンテキストをアクティブ化ルールとして指定し、サポート対象デバイスを指定します。


## <a name="6-install-and-test-the-add-in"></a>6. アドインをインストールおよびテストする

アドインのマニフェスト ファイルで指定した Web サーバーに、HTML ファイル、JavaScript ファイル、CSS ファイルを配置します。アドインをインストールする手順は、アドインの種類によって異なります。詳細については、「[テスト用に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」を参照してください。

Outlook アドインの場合、Exchange メールボックスにインストールし、Exchange 管理センター (EAC) でアドインのマニフェスト ファイルの場所を指定します。詳細については、「[テスト用に Outlook アドインを展開してインストールする](/outlook/add-ins/testing-and-tips)」を参照してください。


## <a name="7-publish-the-add-in"></a>7. アドインを発行する

アドインを AppSource に送信できます。お客様はそこからアドインをインストールできます。さらに、作業ウィンドウおよびコンテンツのアドインを SharePoint 上のプライベート フォルダー アドイン カタログまたは共有ネットワーク フォルダーに発行することが可能で、組織の Exchange サーバーに Outlook アドインを直接展開できます。詳細については、「[Office アドインを発行する](../publish/publish.md)」を参照してください。


## <a name="8-maintain-the-add-in"></a>8. アドインをメンテナンスする

アドインから Web サービスを呼び出していて、アドインの公開後に Web サービスを更新する場合、アドインを再発行する必要はありません。 ただし、アドイン マニフェスト、スクリーンショット、アイコン、HTML、JavaScript のファイルなど、アドインに送信したアイテムやデータを変更する場合は、アドインを再発行する必要があります。 

具体的には、AppSource にアドインを発行した場合は、AppSource が変更を実装できるようにアドインを再送信する必要があります。アドインと一緒に、新しいバージョン番号を含む更新されたアドイン マニフェストを再送信する必要があります。また、新しいマニフェストのバージョン番号と一致するように、送信フォームのアドイン バージョン番号を更新する必要があります。Outlook アドインの場合は、[ID](/office/dev/add-ins/reference/manifest/id) 要素にアドイン マニフェストの異なる UUID が含まれることを確認する必要があります。

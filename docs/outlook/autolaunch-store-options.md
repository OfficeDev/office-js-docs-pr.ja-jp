---
title: イベント ベースの Outlook アドインの AppSource 一覧表示オプション
description: イベント ベースのアクティブ化を実装する Outlook アドインで使用できる AppSource リスト オプションについて説明します。
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d8d2c2e9960d2aef2d32ede6e20eb5f1db125a6c
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797681"
---
# <a name="appsource-listing-options-for-your-event-based-outlook-add-in"></a>イベント ベースの Outlook アドインの AppSource 一覧表示オプション

現時点では、エンド ユーザーがイベント ベースの機能にアクセスするには、組織の管理者がアドインを展開する必要があります。 エンド ユーザーが AppSource から直接アドインを取得した場合、イベント ベースのアクティブ化を制限しています。 たとえば、Contoso アドインにノードの下`LaunchEvents`に少なくとも 1 つ定義`LaunchEvent Type`されている拡張ポイントが含まれている`LaunchEvent`場合、アドインの自動呼び出しは、組織の管理者によってエンド ユーザーに対してアドインがインストールされた場合にのみ行われます。それ以外の場合、アドインの自動呼び出しはブロックされます。 アドイン マニフェストの例から次の抜粋を参照してください。

```xml
...
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    ...
```

エンド ユーザーまたは管理者は、AppSource またはアプリ内 Office ストアを使用してアドインを取得できます。 アドインの主なシナリオまたはワークフローでイベント ベースのアクティブ化が必要な場合は、管理者のデプロイで使用できるアドインを制限することができます。 この制限を有効にするために、フライト コードの URL を指定できます。 フライト コードのおかげで、これらの特別な URL を持つエンド ユーザーのみがリストにアクセスできます。 URL の例を次に示します。

`https://appsource.microsoft.com/product/office/WA200002862?flightCodes=EventBasedTest1`

ユーザーと管理者は、フライト コードが有効になっている場合、AppSource またはアプリ内 Office ストアでその名前でアドインを明示的に検索することはできません。 アドイン作成者は、これらのフライト コードを組織の管理者と非公開で共有し、アドインの展開を行うことができます。

> [!NOTE]
> エンド ユーザーはフライト コードを使用してアドインをインストールできますが、アドインにはイベント ベースのアクティブ化は含まれません。

## <a name="specify-a-flight-code"></a>フライト コードを指定する

アドインに必要なフライト コードを指定するには、アドインを発行するときに、その情報を **Notes の認定用** に共有します。 _**重要**:_ フライト コードでは大文字と小文字が区別されます。

![発行プロセス中の[Notes for Certification] 画面のフライト コードの要求例を示すスクリーンショット。](../images/outlook-publish-notes-for-certification-1.png)

## <a name="deploy-add-in-with-flight-code"></a>フライト コードを使用してアドインをデプロイする

フライト コードが設定されると、アプリ認定チームから URL を受け取ります。 その後、URL を管理者と非公開で共有できます。

アドインをデプロイするには、管理者は次の手順を使用できます。

- Microsoft 365 管理者アカウントで admin.microsoft.com または AppSource.com にサインインします。 アドインでシングル サインオン (SSO) が有効になっている場合は、グローバル管理者の資格情報が必要です。
- フライト コード URL を Web ブラウザーで開きます。
- アドインの一覧ページで、[ **今すぐ入手**] を選択します。 統合アプリ ポータルにリダイレクトする必要があります。

## <a name="unrestricted-appsource-listing"></a>制限のない AppSource の一覧

アドインが重要なシナリオに対してイベント ベースのアクティブ化を使用しない場合 (つまり、アドインが自動呼び出しなしで正常に動作する) 場合は、特別なフライト コードなしで AppSource にアドインを一覧表示することを検討してください。 エンド ユーザーが AppSource からアドインを取得した場合、そのユーザーに対して自動アクティブ化は行われません。 ただし、作業ウィンドウや関数コマンドなど、アドインの他のコンポーネントを使用することもできます。

> [!IMPORTANT]
> これは一時的な制限です。 今後、アドインを直接取得するエンド ユーザーに対して、イベント ベースのアドインのアクティブ化を有効にする予定です。

## <a name="update-existing-add-ins-to-include-event-based-activation"></a>既存のアドインを更新して、イベント ベースのアクティブ化を含める

既存のアドインを更新してイベント ベースのアクティブ化を含め、検証のために再送信し、制限付き AppSource の一覧または制限のない AppSource の一覧を作成するかどうかを決定できます。

更新されたアドインが承認されると、以前にアドインを展開した組織の管理者は、管理センターの **[統合アプリ** ] セクションに更新メッセージを受け取ります。 このメッセージは、イベント ベースのアクティブ化の変更について管理者に通知します。 管理者が変更を受け入れた後、更新プログラムはエンド ユーザーに展開されます。

![[統合アプリ] 画面のアプリ更新通知のスクリーンショット。](../images/outlook-deploy-update-notification.png)

アドインを自分でインストールしたエンド ユーザーの場合、アドインが更新された後でも、イベント ベースのアクティブ化機能は機能しません。

## <a name="admin-consent-for-installing-event-based-add-ins"></a>イベント ベースのアドインをインストールするための管理同意

**統合アプリ** 画面からイベント ベースのアドインが展開されるたびに、管理者は展開ウィザードでアドインのイベント ベースのアクティブ化機能に関する詳細を取得します。 詳細は、[アプリの **アクセス許可と機能] セクションに** 表示されます。 管理者には、アドインが自動的にアクティブ化できるすべてのイベントが表示されます。

![新しいアプリをデプロイするときの [アクセス許可要求を受け入れる] 画面のスクリーンショット。](../images/outlook-deploy-accept-permissions-requests.png)

同様に、既存のアドインがイベント ベースの機能に更新されると、管理者はアドインに "Update Pending" 状態を表示します。 更新されたアドインは、アドインが自動的にアクティブ化できる一連のイベントを含め、[ **アプリのアクセス許可と機能]** セクションに記載されている変更に管理者が同意した場合にのみ展開されます。

アドインに新規 `LaunchEvent Type` を追加するたびに、管理者は管理者ポータルに更新フローを表示し、追加のイベントに同意する必要があります。

![更新されたアプリをデプロイするときの "更新" フローのスクリーンショット。](../images/outlook-deploy-update-flow.png)

## <a name="see-also"></a>関連項目

- [イベント ベースのアクティブ化のために Outlook アドインを構成する](autolaunch.md)

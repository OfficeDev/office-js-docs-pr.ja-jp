---
title: Outlook アドインに関するプライバシー、アクセス許可、セキュリティ
description: Outlook アドインで、プライバシー、アクセス許可、セキュリティを管理する方法について説明します。
ms.date: 10/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 560c9bbdfcde849b66d86e9c000d78f094b3e561
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541257"
---
# <a name="privacy-permissions-and-security-for-outlook-add-ins"></a>Outlook アドインに関するプライバシー、アクセス許可、セキュリティ

エンドユーザー、開発者、および管理者は、Outlook アドインのセキュリティ モデルの階層化されたアクセス許可レベルを使用して、プライバシーとパフォーマンスを制御することができます。

この記事では、Outlook アドインで要求可能なアクセス許可について説明し、次のような観点からセキュリティ モデルを調べます。

- **AppSource**: アドインの整合性

- **エンド ユーザー**: プライバシーとパフォーマンスの問題

- **開発者**: アクセス許可の選択とリソース使用量の制限

- **管理者**: パフォーマンスのしきい値を設定する特権

## <a name="permissions-model"></a>アクセス許可モデル

Because customers' perception of add-in security can affect add-in adoption, Outlook add-in security relies on a tiered permissions model. An Outlook add-in would disclose the level of permissions it needs, identifying the possible access and actions that the add-in can make on the customer's mailbox data.

アクセス許可には、4 つのレベルがあります。

[!include[Table of Outlook permissions](../includes/outlook-permission-levels-table.md)]

アクセス許可の 4 つのレベルは累積的です。**メールボックス読み取り/書き込み** アクセス許可には **アイテム読み取り/書き込み**、**アイテム読み取り**、および **制限付き** が含まれており、**アイテム読み取り/書き込み** には **アイテム読み取り** と **制限付き** が含まれており、また **アイテム読み取り** アクセス許可には **制限付き** が含まれています。

次の図は、アクセス許可の 4 つのレベルを示しています。また、各層でエンド ユーザー、開発者、および管理者に提供される機能が示されています。 これらのアクセス許可の詳細については、「[エンド ユーザー: プライバシーとパフォーマンスについて](#end-users-privacy-and-performance-concerns)」、「[開発者: アクセス許可の選択とリソース使用の制限](#developers-permission-choices-and-resource-usage-limits)」、および「[Outlook アドインのアクセス許可について](understanding-outlook-add-in-permissions.md)」を参照してください。

**4 層のアクセス許可モデルとエンド ユーザー、開発者、および管理者の関連性**

![メール アプリ スキーマ v1.1 の 4 層アクセス許可モデルの図。](../images/add-in-permission-tiers.png)

## <a name="appsource-add-in-integrity"></a>AppSource: アドインの整合性

[AppSource](https://appsource.microsoft.com) は、エンド ユーザーと管理者がインストールできるアドインをホストします。 AppSource は、これらの Outlook アドインの整合性を維持するために次の手段を適用します。

- アドインのホスト サーバーは必ず Secure Socket Layer (SSL) を使用して通信する必要があります。

- 開発者はアドインを提出する際に、ID の証明、契約上の合意、および法規制に準拠したプライバシー ポリシーを提供する必要があります。

- アドインを読み取り専用モードでアーカイブします。

- 使用可能なアドインに対するユーザーレビュー システムをサポートしてコミュニティの自己管理を促します。

## <a name="optional-connected-experiences"></a>オプションの接続エクスペリエンス

エンド ユーザーと IT 管理者は、[Office のデスクトップ クライアントとモバイル クライアントでオプションの接続エクスペリエンスを](/deployoffice/privacy/optional-connected-experiences) オフにすることができます。 Outlook アドインの場合、 **オプションの接続エクスペリエンス** 設定を無効にした場合の影響はクライアントによって異なりますが、通常は、ユーザーがインストールしたアドインと Office ストアへのアクセスが許可されないことを意味します。 [一元展開](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)を通じて組織の IT 管理者が展開したアドインは引き続き使用できます。

- Windows\*、Mac: [ **アドインの取得]** ボタンが表示されないので、ユーザーはアドインを管理したり、Office ストアにアクセスしたりできなくなります。
- Android、iOS: **[アドインの取得]** ダイアログには、管理者が展開したアドインのみが表示されます。
- ブラウザー: アドインの可用性とストアへのアクセスは影響を受けないため、ユーザーは [アドイン （管理者が展開したものを含む） を引き続き管理](https://support.microsoft.com/office/8f2ce816-5df4-44a5-958c-f7f9d6dabdce) できます。

  > [!NOTE]
  > \* Windows の場合、このエクスペリエンス/動作のサポートはバージョン 2008 (ビルド 13127.20296) から入手できます。 バージョンに応じた詳細については、[Microsoft 365](/officeupdates/update-history-office365-proplus-by-date)更新履歴ペーのページと、[Office クライアントのバージョンを見つけてチャネルを更新する方法](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)を参照してください。

アドインの全般的な動作については、「[Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md#optional-connected-experiences)」を参照してください。

## <a name="end-users-privacy-and-performance-concerns"></a>エンド ユーザー: プライバシーとパフォーマンスの問題

セキュリティ モデルによって、エンド ユーザーのセキュリティ、プライバシー、およびパフォーマンスの問題に次のような方法で対処します。

- Outlook の Information Rights Management (IRM) によって保護されているエンド ユーザーのメッセージは、Windows 以外のクライアントで Outlook アドインと対話しません。

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Before installing an add-in from AppSource, end users can see the access and actions that the add-in can make on their data and must explicitly confirm to proceed. No Outlook add-in is automatically pushed onto a client computer without manual validation by the user or administrator.

- **制限付き** のアクセス許可を付与すると、Outlook アドインは現在のアイテムのみに限定されたアクセスが可能になります。 **アイテムの読み取り** アクセス許可を付与すると、Outlook アドインは現在のアイテムに対してのみ、送信者と受信者の名前や電子メール アドレスなどの個人を特定できる情報にアクセスできます。

- An end user can install an Outlook add-in for only himself or herself. Outlook add-ins that affect an organization are installed by an administrator.

- エンド ユーザーは、ユーザーのセキュリティ リスクを最小限に抑えながら、ユーザーにとって魅力的な状況依存のシナリオを実現する Outlook アドインをインストールできます。

- インストールされた Outlook アドインのマニフェスト ファイルは、ユーザーの電子メール アカウントに安全に保管されます。

- Office アドインをホストするサーバーと通信するデータは、Secure Socket Layer (SSL) プロトコルで常に暗号化されます。

- Outlook リッチ クライアントのみ: Outlook リッチ クライアントは、インストールされた Outlook アドインのパフォーマンスを監視し、ガバナンス制御を実施し、次の領域で制限を超えている Outlook アドインを無効にします。

  - アクティブ化までの応答時間

  - アクティブ化または再アクティブ化に失敗した回数

  - メモリ使用量

  - CPU 使用率  

  Governance deters denial-of-service attacks and maintains add-in performance at a reasonable level. The Business Bar alerts end users about Outlook add-ins that the Outlook rich client has disabled based on such governance control.

- エンド ユーザーは、いつでも Exchange 管理センターで、インストールした Outlook アドインから要求されたアクセス許可を確認したり、Outlook アドインを無効にしたり、その後で有効にしたりできます。

## <a name="developers-permission-choices-and-resource-usage-limits"></a>開発者: アクセス許可の選択とリソース使用量の制限

開発者は、セキュリティ モデルで規定されたきめ細かいレベルのアクセス許可を選択し、厳密なパフォーマンス ガイドラインを守る必要があります。

### <a name="tiered-permissions-increases-transparency"></a>階層化された許可で透過性が向上

開発者は階層化された許可モデルに従うことにより、透明性を提供しつつ、アドインがデータとメールボックスに対して実行可能なアクションに対するユーザーの懸念を緩和し、アドインの導入を間接的に促進できます。

- 開発者は、Outlook アドインがアクティブ化される方法、およびメール アドインがアイテムの特定のプロパティを読み書きする必要性や、アイテムを作成および送信する必要性に基づいて、Outlook アドインの適切なレベルのアクセス許可を要求します。

- 前述のように、開発者はマニフェストでアクセス許可を要求します。

  次の例では、XML マニフェストで **読み取り項目** のアクセス許可を要求します。

  ```XML
  <Permissions>ReadItem</Permissions>
  ```

  次の例では、Teams マニフェスト (プレビュー) で **アイテムの読み取り** アクセス許可を要求します。

```json
"authorization": {
  "permissions": {
    "resourceSpecific": [
      ...
      {
        "name": "MailboxItem.Read.User",
        "type": "Delegated"
      },
    ]
  }
},
```

- 開発者は、Outlook アドインが特定の種類の Outlook アイテム (予定またはメッセージ) またはアイテムの件名または本文に存在する特定の抽出されたエンティティ (電話番号、住所、URL) でアクティブ化された場合、 **制限付** きアクセス許可を要求できます。 たとえば、次のルールは、現在のメッセージの件名または本文に電話番号、郵送先住所、URL の 3 つのエンティティのうち 1 つ以上のエンティティが見つかった場合に Outlook アドインをアクティブ化します。

> [!NOTE]
> この例に示すように、ライセンス認証規則は、 [Office アドインの Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用するアドインではサポートされていません。

  ```XML
    <Permissions>Restricted</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        </Rule>
    </Rule>
  ```

- 開発者は、Outlook アドインが既定の抽出されたエンティティ以外の現在のアイテムのプロパティを読み取る必要がある場合、または現在のアイテムのアドインによって設定されたカスタム プロパティを書き込む必要があるが、他のアイテムの読み取りまたは書き込み、またはユーザーのメールボックスでのメッセージの作成または送信を必要としない場合は、アイテムの読み取 **りアクセス許可** を要求する必要があります。 たとえば、Outlook アドインでアイテムの件名または本文に含まれる会議開催の提案、タスクの提案、メール アドレス、連絡先名などのエンティティを検索する必要がある場合や、アクティブ化に正規表現を使用する必要がある場合は、**アイテムの読み取り** のアクセス許可を要求します。

- Outlook アドインが新規作成アイテムのプロパティ (受信者名、メールアドレス、本文、件名など) を書き込む必要がある場合、またはアイテムの添付ファイルを追加または削除する必要がある場合、開発者は **アイテムの読み取り/書き込み** 許可を要求します。

- 開発者は、Outlook アドインで [mailbox.makeEWSRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して次のいずれか 1 つ以上の処理を実行する必要がある場合にのみ、**メールボックスの読み取り/書き込み** のアクセス許可を要求します。

  - メールボックスのアイテムのプロパティに対する読み取りまたは書き込み。
  - メールボックスのアイテムの作成、読み取り、書き込み、または送信。
  - メールボックスのフォルダーの作成、読み取り、または書き込み。

### <a name="resource-usage-tuning"></a>リソース使用量の調整

Developers should be aware of resource usage limits for activation, incorporate performance tuning in their development workflow, so as to reduce the chance of a poorly performing add-in denying service of the host. Developers should follow the guidelines in designing activation rules as described in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md). If an Outlook add-in is intended to run on an Outlook rich client, then developers should verify that the add-in performs within the resource usage limits.

### <a name="other-measures-to-promote-user-security"></a>ユーザーのセキュリティを高めるその他の方法

開発者は、以下の点についても意識し、計画する必要があります。

- 開発者は、サポートされていないため、アドインで ActiveX コントロールを使用できません。

- 開発者は AppSource に Outlook アドインを提出する際に、次の作業を実行する必要があります。

  - ID の証明として Extended Validation (EV) SSL 証明書を生成する。

  - SSL をサポートする Web サーバーで、提出するアドインをホストする。

  - 準拠したプライバシー ポリシーを生成する。

  - アドインの提出時に契約合意書に署名する。

## <a name="administrators-privileges"></a>管理者: 特権

セキュリティ モデルによって、管理者に次の権利と責任が与えられます。

- AppSource のアドインを含めて、エンド ユーザーが Outlook アドインをインストールできないようにすることができます。

- Exchange 管理センターで Outlook アドインを無効または有効にできます。

- Windows 版 Outlook のみ: GPO レジストリ設定を使用して、パフォーマンスのしきい値の設定を無効にすることができます。

## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
- [Microsoft 365 アプリのプライバシー コントロール](/deployoffice/privacy/overview-privacy-controls)
- [Outlook アドインの API](apis.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)

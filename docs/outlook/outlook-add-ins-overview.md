---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 08/18/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 83644823f4ca906f52cae430fa3a7f350dbf076c
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803780"
---
# <a name="outlook-add-ins-overview"></a>Outlook アドインの概要

Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。 Outlook アドインには次の 3 つの主な側面があります。

- Windows と Mac 用のデスクトップ Outlook、Web 版 (Microsoft 365 と Outlook.com)、モバイル版すべてで機能する同じアドインとビジネス ロジック。
- Outlook アドインは、マニフェスト (アドインが Outlook に統合する方法 (ボタンや作業ウィンドウなど) を説明する)、および JavaScript/HTML のコード (アドインの UI とビジネス ロジックを構成する) で構成される。
- Outlook アドインは、[AppSource](https://appsource.microsoft.com) から入手するか、エンドユーザーまたは管理者が[サイドロード](sideload-outlook-add-ins-for-testing.md)することができます。

Outlook アドインは、Windows 版 Outlook 固有の統合機能として以前から存在した COM アドインや VSTO アドインとは異なります。 COM アドインとは違い、Outlook アドインのコードがユーザーのデバイスまたは Outlook クライアントに物理的にインストールされることはありません。 Outlook のアドインの場合、Outlook はマニフェストを読み取り UI で指定したコントロールをフックした後に、HTML と JavaScript を読み込みます。 この Web コンポーネントは、サンドボックス内のブラウザーのコンテキストですべて実行されます。

アドインをサポートしている Outlook アイテムには、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消し、予定などがあります。 それぞれの Outlook アドインにより、アイテムの種類、ユーザーがアイテムの読み取りや作成を行うかどうかなど、使用できるコンテキストが定義されます。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>拡張点

拡張点は、アドインが Outlook と統合する方法です。これを行う方法は以下のとおりです。

- アドインは、メッセージと予定のコマンド サーフェスに表示されるボタンを宣言できます。詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。

    **リボン上の [コマンド] ボタンがあるアドイン**

    ![アドイン コマンドの UI なし図形](../images/uiless-command-shape.png)

- アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。

    **強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**

    ![カード内のコンテキスト アプリを示しています](../images/outlook-detected-entity-card.png)

> [!NOTE]
> [カスタム ウィンドウは廃止された](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)ため、サポートされている拡張点を使用していることを確認してください。

## <a name="mailbox-items-available-to-add-ins"></a>アドインで使用可能なメールボックスのアイテム

Outlook アドインは、作成中や読み取り中にメッセージや予定で使用することができますが、他のアイテムの種類では使用できません。新規作成フォームまたは閲覧フォームで現在のメッセージ アイテムが次のいずれかの場合、Outlook はアドインをアクティブ化しません。

- Information Rights Management (IRM) によって保護されているか、または保護のためにその他の方法で暗号化されている場合。デジタル署名はこれらいずれかのメカニズムに依存しているため、デジタル署名されたメッセージはその一例です。

  > [!IMPORTANT]
  > - アドインは、Microsoft 365 サブスクリプションに関連付けられている Outlook のデジタル署名付きメッセージでライセンス認証を行います。 Windows では、このサポートはビルド 8711.1000 で導入されました。
  >
  > - Windows の Outlook ビルド13120.1000 から、アドインは IRM で保護されたアイテムに対して有効になるようになりました。 この機能のプレビューの詳細については、「[Information Rights Management (IRM) で保護されているアイテムのアドインのアクティブ化](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)」を参照してください。

- メッセージ クラスが IPM.Report.* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。

- 下書きであるか (送信者が割り当てられていない)、Outlook の [下書き] フォルダーにある場合。

- 別のメッセージに添付される .msg または .eml ファイルの場合。

- .msg または .eml ファイルがファイル システムから開かれた場合。

- 共有メールボックス内、別のユーザーのメールボックス内、アーカイブ メールボックス内、パブリック フォルダー内。

- カスタム フォームを使用する場合。

既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。 この理由の詳細は、[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)の「既知のエンティティに対するサポート」をご覧ください。

## <a name="supported-hosts"></a>サポートされるホスト

Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Outlook on the web と Outlook.com でサポートされています。 最新の機能すべてが、すべての[クライアント](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。 これらの機能が各ホストでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。


## <a name="get-started-building-outlook-add-ins"></a>Outlook アドインの作成を開始する

Outlook アドインの作成を開始するには、次の操作を行います。

- [クイックスタート](../quickstarts/outlook-quickstart.md) - 簡単な作業ウィンドウを作成します。
- [チュートリアル](../tutorials/outlook-tutorial.md) - 新しいメッセージに GitHub gist を挿入するアドインを作成する方法について説明します。


## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [Office および SharePoint アドインのライセンスを付与する](/office/dev/store/license-your-add-ins)
- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-the-office-store)

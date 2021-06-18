---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 06/15/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: f0c1dbdd1cf9909310b629188d4f3d3d5de6b6bb
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007812"
---
# <a name="outlook-add-ins-overview"></a>Outlook アドインの概要

Outlook アドインは、Web ベースのプラットフォームを使用してサードパーティ企業によって Outlook に組み込まれた統合機能です。Outlook アドインには次の 3 つの主な側面があります。

- Windows と Mac 用のデスクトップ Outlook、Web 版 (Microsoft 365 と Outlook.com)、モバイル版すべてで機能する同じアドインとビジネス ロジック。
- Outlook アドインは、マニフェスト (アドインが Outlook に統合する方法 (ボタンや作業ウィンドウなど) を説明する)、および JavaScript/HTML のコード (アドインの UI とビジネス ロジックを構成する) で構成される。
- Outlook アドインは、[AppSource](https://appsource.microsoft.com) から入手するか、エンドユーザーまたは管理者が[サイドロード](sideload-outlook-add-ins-for-testing.md)することができます。

Outlook アドインは、Windows で実行する Outlook に固有の古い統合である COM アドインや VSTO アドインとは異なります。COM アドインとは異なり、Outlook アドインには、ユーザーのデバイスや Outlook クライアントに物理的にインストールされたコードがありません。Outlook アドインの場合、Outlook はマニフェストを読み取り、指定された UI コントロールをフックして、JavaScript と HTML を読み込みます。Web コンポーネントは全て、サンドボックス内のブラウザーのコンテキストで実行されます。

アドインをサポートする Outlook アイテムには、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消し、予定などがあります。それぞれの Outlook アドインでは、メール アドインが使用できるコンテキストを定義します。これにはアイテムの種類、およびユーザーがアイテムの読み取り (または作成) を行っているかどうかなどがあります。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>拡張点

拡張点は、アドインが Outlook と統合する方法です。これを行う方法は以下のとおりです。

- アドインは、メッセージと予定のコマンド サーフェスに表示されるボタンを宣言できます。詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。

    **リボン上の [コマンド] ボタンがあるアドイン**

    ![アドイン コマンドの UI なし図形](../images/uiless-command-shape.png)

- アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。

    **強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**

    ![カード内のコンテキスト アプリを示す](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>アドインで使用可能なメールボックスのアイテム

Outlook アドインは、ユーザーがメッセージまたは予定を作成または読んでいるときにアクティブになりますが、他の種類のアイテムではアクティブになりません。 ただし、現在のメッセージ アイテムが作成または読み取りフォームで次のいずれかである場合、アドインはアクティブ化 *されません*。

- Information Rights Management (IRM) によって保護されているか、または保護のためにその他の方法で暗号化されている場合。デジタル署名はこれらいずれかのメカニズムに依存しているため、デジタル署名されたメッセージはその一例です。

  > [!IMPORTANT]
  >
  > - アドインは、Microsoft 365 サブスクリプションに関連付けられている Outlook のデジタル署名付きメッセージでライセンス認証を行います。 Windows では、このサポートはビルド 8711.1000 で導入されました。
  >
  > - Windows の Outlook ビルド 13229.10000 から、IRM で保護されたアイテムに対してアドインをアクティブ化できるようになりました。 この機能のプレビューの詳細については、「[Information Rights Management (IRM) で保護されているアイテムのアドインのアクティブ化](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)」を参照してください。

- メッセージ クラスが IPM.Report.* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。

- 別のメッセージに添付される .msg または .eml ファイルの場合。

- .msg または .eml ファイルがファイル システムから開かれた場合。

- 共有メールボックス\*の[グループ メールボックス](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)内、別のユーザーのメールボックス内\*、アーカイブ メールボックス内、パブリック フォルダー内。

  > [!IMPORTANT]
  > \* [要件セット 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) では、代理アクセス シナリオ (別のユーザーのメールボックスで共有されるフォルダなど) のサポートが導入されました。 共有メールボックスのサポートをプレビューしています。 詳細については、「[共有フォルダーと共有メールボックスのシナリオを有効にする](delegate-access.md)」を参照してください。

- カスタム フォームを使用する場合。

既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。 この理由の詳細は、[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)の「既知のエンティティに対するサポート」をご覧ください。

## <a name="supported-clients"></a>サポートされるクライアント

Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Outlook on the web と Outlook.com でサポートされています。 最新の機能すべてが、すべての[クライアント](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。 これらの機能が各アプリケーションでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。

## <a name="get-started-building-outlook-add-ins"></a>Outlook アドインの作成を開始する

Outlook アドインの作成を開始するには、次の操作を行います。

- [クイックスタート](../quickstarts/outlook-quickstart.md) - 簡単な作業ウィンドウを作成します。
- [チュートリアル](../tutorials/outlook-tutorial.md) - 新しいメッセージに GitHub gist を挿入するアドインを作成する方法について説明します。

## <a name="see-also"></a>関連項目

- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)
- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [Office アドインの設計ガイドライン](../design/add-in-design.md)
- [Office および SharePoint アドインのライセンスを付与する](/office/dev/store/license-your-add-ins)
- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-the-office-store)

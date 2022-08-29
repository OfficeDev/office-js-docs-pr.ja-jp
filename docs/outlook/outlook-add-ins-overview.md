---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 08/09/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 0503a0cfae39e58c11fefc6cc87a239d7ecdbc05
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2022
ms.locfileid: "67318813"
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

    ![アドイン関数コマンド。](../images/uiless-command-shape.png)

- アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。

    **強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**

    ![カード内のコンテキスト アプリを示す。](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>アドインで使用可能なメールボックスのアイテム

Outlook アドインは、ユーザーがメッセージまたは予定を作成または読んでいるときにアクティブになりますが、他の種類のアイテムではアクティブになりません。 ただし、現在のメッセージ アイテムが作成または読み取りフォームで次のいずれかである場合、アドインはアクティブ化 *されません*。

- Information Rights Management (IRM) によって保護されるか、保護のために他の方法で暗号化され、Windows 以外のクライアント上の Outlook からアクセスされます。 デジタル署名されたメッセージはその一例で、デジタル署名はこれらのメカニズムのどちらかに依存しています。

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- メッセージ クラスが IPM.Report.* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。

- 別のメッセージに添付される .msg または .eml ファイルの場合。

- .msg または .eml ファイルがファイル システムから開かれた場合。

- 共有メールボックス\*の[グループ メールボックス](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)内、別のユーザーのメールボックス\*内、[アーカイブ メールボックス](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-features#archive-mailbox)内、パブリック フォルダー内。

  > [!IMPORTANT]
  > \* [要件セット 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) では、代理アクセス シナリオ (別のユーザーのメールボックスで共有されるフォルダなど) のサポートが導入されました。 共有メールボックスのサポートは、Windowsお よび Mac の Outlook でプレビューになりました。 詳細については、「 [共有フォルダーと共有メールボックスを有効にするシナリオ](delegate-access.md)」を参照してください。

- カスタム フォームを使用する場合。

- 簡易 MAPI で作成されます。 簡易 MAPI は、Outlook が閉じられている間に Office ユーザーが Windows 上の Office アプリケーションからメールを作成または送信するときに使用されます。 たとえば、ユーザーは Word での作業中に Outlook メールを作成できます。これにより、Outlook アプリケーション全体を起動せずに Outlook メール作成ウィンドウがトリガーされます。 ただし、ユーザーが Word からメールを作成するときに Outlook が既に実行されている場合、これは簡易 MAPI シナリオではないため、Outlook アドインは、他のアクティブ化要件が満たされている限り、作成フォームで動作します。

既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。 この背後にある理由の詳細については、「[既知のエンティティに対するサポート](match-strings-in-an-item-as-well-known-entities.md#support-for-well-known-entities)」をご覧ください。

現在、モバイル クライアント用のアドインを設計および実装する際には、さらに考慮事項があります。 詳細については、「 [モバイル サポートを Outlook アドインに追加する](add-mobile-support.md#compose-mode-and-appointments)」を参照してください。

## <a name="supported-clients"></a>サポートされるクライアント

Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Outlook on the web と Outlook.com でサポートされています。 最新の機能すべてが、すべての[クライアント](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。 これらの機能が各アプリケーションでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。

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

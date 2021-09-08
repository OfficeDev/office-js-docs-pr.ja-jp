---
title: iPad アドインの特別な要件
description: カスタム アドインで実行される Officeアドインを作成するための要件について説明iPad。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 04ee1a4bea8b9f27189bf67368f883cd3b91ab3e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938963"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>iPad アドインの特別な要件

アドインでサポートされている API Officeのみを使用する場合は、iPadを iPad にインストールできます。 (詳細 [についてはOffice API 要件を指定する](specify-office-hosts-and-api-requirements.md)を参照してください。*アドインが [AppSource](https://appsource.microsoft.com)* を通じて販売される場合は、すべての [Office](../concepts/add-in-development-best-practices.md)アドインに適用されるベスト プラクティスに加えて、iPad にインストールできるアドインに対して従う必要があるプラクティスがあります。

次の表に、実行するタスクの一覧を示します。

> [!NOTE]
> Outlook Mobile でOutlook良く動作するアドインを設計する方法については、「Outlook Mobile 用アドイン」を[参照してください](../outlook/outlook-mobile-addins.md)。

|タスク|説明|リソース|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|iOS デザインのベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。| 以下のメモを参照してください。 |
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[認定ポリシー 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|アドインのコマースをオンラインで無料iPad。|iPad で実行する場合、アドインには、アプリ内購入、試用版の提供、無料版へのアップセルを目的とした UI、またはユーザーが他のコンテンツ、アプリ、アドインを購入または取得できるオンライン ストアへのリンクが含されていないことが必要です。また、プライバシー ポリシーと利用規約ページには、コマース UI または AppSource リンクが含されていないことも必要です。|[認定ポリシー 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>アドインは、他のプラットフォームでコマースを持ち続け得る。 これを行うには[、Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed)プロパティをテストし、返す際にすべての商取引を抑制します `false` 。|
|アドインを AppSource に送信します。|パートナー センターの [製品のセットアップ] ページで **、[iOS** と Android で製品を利用可能にする (該当する場合) ] チェック ボックスをオンにし、[アカウント設定] で Apple 開発者 ID を入力します。 アプリケーション プロバイダー [契約を確認](https://go.microsoft.com/fwlink/?linkid=715691) して、条件を理解してください。|[AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> アドインは、実行中のデバイスに基づいて代替 UI を提供できます。 アドインがアプリで実行されているかどうかを検出iPad、次の API を使用できます。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchEnabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceAllowed)
>
> 次のiPad `touchEnabled` を `true` `commerceAllowed` 返します `false` 。
>
> アプリの最適な UI 設計方法については、「iPad iOS 用のデザイン」[を参照してください](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>アプリで実行できるOfficeアドインを開発するためのベスト プラクティスiPad

アプリで実行するアドインを開発するには、次のベスト プラクティスをiPad。

-  **アドインを開発およびデバッグするには、Windowsまたは Mac でアドインをサイドロードiPad。**

    iPad でアドインを直接開発することはできませんが、Windows または Mac コンピューターで開発およびデバッグし、テスト用に iPad にサイドロードできます。 iOS または Mac の Office で実行されるアドインは、Windows の Office で実行されているアドインと同じ API をサポートしています。そのため、アドインのコードは、これらのプラットフォームで同じ方法で実行する必要があります。 詳細については[、「Test and debug Office](../testing/test-debug-office-add-ins.md)アドイン」および[「Office iPad](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)Mac のサイドロード アドイン」を参照してください。

-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**

    アドインのマニフェストで API 要件を指定すると、Office クライアント アプリケーションがそれらの API メンバーをサポートOffice判断します。 API メンバーがアプリケーションで使用できる場合は、アドインを使用できます。 または、ランタイム チェックを実行して、アドインでメソッドを使用する前に、アプリケーションでメソッドを使用できるかどうかを判断することもできます。 ランタイム チェックでは、アドインがアプリケーションで常に使用可能であり、メソッドが使用可能な場合は追加の機能が提供されます。 詳細については、「アプリケーションと[API の要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。

---
title: iPad アドインの特別な要件
description: iPad で実行されるカスタム アドインOffice作成するためのいくつかの要件について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: fdb402f4302e7e81589d586fa1ecd5b30d4e515d
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237855"
---
# <a name="special-requirements-for-add-ins-on-the-ipad"></a>iPad アドインの特別な要件

アドインで iPad でOfficeされている API のみを使用する場合、ユーザーは iPad にインストールできます。 (詳細 [についてはOfficeアプリケーションと API の要件を指定する](specify-office-hosts-and-api-requirements.md) を参照してください)。 *アドインが [AppSource](https://appsource.microsoft.com)* を通じて販売される場合は、すべての [Office](../concepts/add-in-development-best-practices.md)アドインに適用されるベスト プラクティスに加えて、iPad にインストールできるアドインに関して従う必要があるいくつかのプラクティスがあります。

次の表に、実行するタスクを示します。

> [!NOTE]
> 外観が良く、Outlook Mobile でうまく機能する Outlook アドインの設計については [、「Outlook Mobile](../outlook/outlook-mobile-addins.md)用アドイン」を参照してください。

|Task|説明|リソース|
|:-----|:-----|:-----|
|アドインを更新して、Office.js バージョン 1.1 をサポートします。|Office アドイン プロジェクトで使用する JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。|[API とマニフェストのバージョンを更新する](update-your-javascript-api-for-office-and-manifest-schema-version.md)|
|iOS 設計のベスト プラクティスを適用します。|アドイン UI を iOS エクスペリエンスとシームレスに統合します。| 以下の注を参照してください。 |
|タッチ用にアドインを最適化します。|マウスとキーボードに加え、タッチ入力に対して、UI が素早く応答するようにします。|[UX 設計原則を適用する](../concepts/add-in-development-best-practices.md#apply-ux-design-principles)|
|アドインを無料にします。|iPad 上の Office は、ユーザー数を拡大して、サービスを促進できるチャネルです。これらの新しいユーザーは、お客様になる可能性があります。|[認定ポリシー 1120.2](/legal/marketplace/certification-policies#11202-acquisition-pricing-and-terms)|
|iPad でアドインの商取引を無料にする。|iPad で実行する場合、アドインは、アプリ内購入、試用版の提供、非無料バージョンへのアップセルを目的とした UI、またはユーザーが他のコンテンツ、アプリ、またはアドインを購入または取得できる任意のオンライン ストアへのリンクを持っている必要があります。プライバシー ポリシーと使用条件のページには、商取引 UI や AppSource のリンクも含めずに設定する必要があります。|[認定ポリシー 1100.3](/legal/marketplace/certification-policies#11003-selling-additional-features)<br><br>アドインは、他のプラットフォームでも商取引を行います。 これを行うには [、Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed) プロパティをテストし、戻り値が返された場合は、すべての商取引を抑制します `false` 。|
|アドインを AppSource に提出します。|パートナー センターの [製品のセットアップ] ページで **、[iOS** と Android で製品を利用可能にする (該当する場合)] チェック ボックスをオンにし、[アカウント設定] で Apple 開発者 ID を入力します。 アプリケーション プロバイダー [契約を確認](https://go.microsoft.com/fwlink/?linkid=715691) して、条項を理解してください。|[AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-appsource-via-partner-center)|

> [!NOTE]
> アドインは、実行されているデバイスに基づいて代替 UI を提供できます。 アドインが iPad で実行されているかどうかを検出するには、次の API を使用できます。
>
> - var isTouchEnabled = [Office.context.touchEnabled](/javascript/api/office/office.context#touchenabled)
> - var allowCommerce = [Office.context.commerceAllowed](/javascript/api/office/office.context#commerceallowed)
>
> iPad では、戻 `touchEnabled` り値 `true` を `commerceAllowed` 返します `false` 。
>
> iPad の最適な UI 設計プラクティスの詳細については [、「iOS 向け設計」を参照してください](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)。

## <a name="best-practices-for-developing-office-add-ins-that-can-run-on-ipad"></a>iPad で実行Officeアドインを開発するためのベスト プラクティス

iPad で実行されるアドインを開発するには、次のベスト プラクティスを適用します。

-  **Windows または Mac でアドインを開発およびデバッグし、iPad にサイドロードします。**

    アドインを iPad で直接開発することはできませんが、Windows または Mac コンピューターで開発とデバッグを行い、テストのために iPad にサイドロードすることができます。 iOS または Mac 上の Office で実行されるアドインは、Windows 上の Office で実行されているアドインと同じ API をサポートします。そのため、アドインのコードは、これらのプラットフォームで同じ方法で実行する必要があります。 詳細については、「アドイン [のテストとデバッグOffice、](../testing/test-debug-office-add-ins.md) テスト用に iPad Office Mac でアドインをサイドロード [する」を参照してください](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)。

-  **アドインのマニフェストまたはランタイム チェックを使用して API の要件を指定する。**

    アドインのマニフェストで API の要件を指定すると、Office クライアント アプリケーションがそれらの API メンバーをサポートOfficeが決定します。 API メンバーがアプリケーションで使用可能な場合は、アドインが使用可能になります。 または、アドインでメソッドを使用する前に、ランタイム チェックを実行して、メソッドがアプリケーションで使用可能かどうかを判断することもできます。 ランタイム チェックでは、アドインが常にアプリケーションで使用できると確認し、メソッドが使用可能な場合は追加の機能を提供します。 詳細については、「アプリケーションと [API の要件Office指定する」を参照してください](specify-office-hosts-and-api-requirements.md)。

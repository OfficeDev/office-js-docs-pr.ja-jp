---
title: マニフェスト ファイルの IconUrl 要素
description: IconUrl 要素は、挿入 UX とストア内のアドインOfficeを表すイメージの URL をOfficeします。
ms.date: 03/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8f4026c9f97fc147b65532e370e18327bc07076c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154945"
---
# <a name="iconurl-element"></a>IconUrl 要素

挿入 UX と Office ストアで Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができるもの

[Override](override.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>注釈

メール アドインの場合、アイコンは [ファイル管理]アドイン UI (Outlook) または 設定 [アドインの管理  >  ] UI (Outlook on the web)  >  に表示されます。 コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]** > **[アドイン]** UI に表示されます。 すべてのアドインの種類に対して、アドインを AppSource に発行する場合、アイコンは [AppSource](https://appsource.microsoft.com)でも使用されます。

画像のファイル形式は GIF、JPG、PNG、EXIF、BMP、TIFF のいずれかにする必要があります。 コンテンツ アプリおよび作業ウィンドウ アプリの場合、指定する画像は 32 x 32 ピクセルにする必要があります。 メール アプリの場合、画像の解像度は 64 x 64 ピクセルである必要があります。 [HighResolutionIconUrl](highresolutioniconurl.md)要素を使用して、高 DPI Officeで実行されているクライアント アプリケーションで使用するアイコンも指定する必要があります。 詳細については、「[効果的な AppSource と Office 内の登録リストを作成する](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)」の「_アプリに一貫性のあるビジュアル ID を作成する_」セクションを参照してください。

実行時に要素の値 `IconUrl` を変更することはできません。現在はサポートされていません。

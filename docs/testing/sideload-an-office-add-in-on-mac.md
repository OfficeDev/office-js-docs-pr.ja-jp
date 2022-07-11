---
title: テスト用の Mac 上の Office アドインをサイドロードする
description: サイドローディングを使用して、Mac で Office アドインをテストします。
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38ed5f5dba2d379b6137a098240021bd642d6e11
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713219"
---
# <a name="sideload-office-add-ins-on-mac-for-testing"></a>テスト用の Mac 上の Office アドインをサイドロードする

Office on Mac でのアドインの実行方法を確認するには、アドインのマニフェストをサイドロードします。 このアクションでは、実行中、ブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使えることと適切にレンダリングされることを確認できます。

> [!NOTE]
> Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

## <a name="prerequisites-for-office-on-mac"></a>Office on Mac の前提条件

- [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。

- Word on Mac バージョン 15.18 (160109)。

- Excel on Mac バージョン 15.19 (160206)。

- PowerPoint on Mac バージョン 15.24 (160614)。

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="sideload-an-add-in-in-office-on-mac"></a>Office on Mac にアドインをサイドロードする

1. **Finder を** 使用してマニフェスト ファイルをサイドロードします。 **Finder を** 開き、Command + Shift + G と入力して **、[フォルダーに移動**] ダイアログを開きます。

1. サイドローディングに使用するアプリケーションに基づいて、次のいずれかのファイルパスを入力します。 `wef` フォルダーがコンピューター上に存在しない場合は、作成します。

    - Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > 残りの手順では、Word アドインをサイドロードする方法について説明します。

1. アドインのマニフェスト ファイルをこの `wef` フォルダーにコピーします。

    ![Office on Mac の Wef フォルダー。](../images/all-my-files.png)

1. Word を起動し、ドキュメントを開きます。 既に起動している場合は、Word を再起動します。

1. Word で、[アドイン **の** > **挿入** > **] [マイ アドイン**] (ドロップダウン メニュー) を選択し、アドインを選択します。

    ![Office on Mac のアドイン。](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。

1. アドインが Word に表示されることを確認します。

    ![Office アドインが Office on Mac に表示されます。](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

以前にサイドロードされたアドインを削除するには、コンピューター上の Office キャッシュをクリアします。 プラットフォームとアプリケーションごとにキャッシュをクリアする方法の詳細については、「 [Office キャッシュをクリアする」](clear-cache.md)を参照してください。

## <a name="see-also"></a>関連項目

- [iPad でテスト用の Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad.md)
- [Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)

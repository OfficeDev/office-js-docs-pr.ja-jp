---
title: 新規作成フォーム用の Outlook アドインを作成する
description: 新規作成フォーム用の Outlook アドインのシナリオと機能について説明します。
ms.date: 02/09/2021
ms.localizationpriority: high
ms.openlocfilehash: 4dd39bd73cbc1c37783c13adb9b4f9f98925b9cf
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484002"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>新規作成フォーム用の Outlook アドインを作成する

Office アドイン マニフェストのスキーマのバージョン 1.1 と Office.js の v1.1 から、新規作成フォームでアクティブ化される Outlook アドインである、新規作成アドインを作成できるようになりました。閲覧アドイン (ユーザーがメッセージや予定を表示するとき閲覧モードでアクティブ化される Outlook アドイン) とは対照的に、新規作成アドインは以下のユーザー シナリオで使用できます。

- 新しいメッセージ、会議出席依頼または予定を新規作成フォームで作成している。

- 既存の予定またはユーザーが開催者になっている会議アイテムを表示または編集している。

   > [!NOTE]
   > ユーザーが Outlook 2013 および Exchange 2013 の RTM リリースを使用していて、ユーザー別の会議アイテムを表示している場合は、使用可能な閲覧アドインを検索できます。Office 2013 SP1 リリース以降では、同じシナリオにおいて作成アドインのみをアクティブ化して使用できるような変更が行われています。

- インライン応答メッセージを作成しているか、別の新規作成フォームでメッセージに返信している。

- 会議出席依頼または会議アイテムに対する応答 ([**承諾**]、[**仮承諾**]、[**辞退**]) を編集している。

- 会議アイテム用に新しい時間を提案している。

- 会議出席依頼や会議アイテムを転送するか、それらに返信している。

これらの新規作成シナリオのそれぞれで、アドインによって定義されているアドイン コマンド ボタンが表示されます。アドイン コマンドが実装されていない古いアドインの場合は、ユーザーはリボンの **[Office アドイン]** を選択してアドイン選択ウィンドウを開いてから、新規作成アドインを選択して開始することができます。次の図は、新規作成フォームのアドイン コマンドを示しています。

![アドイン コマンドが含まれた Outlook 作成フォームが表示されています。](../images/compose-form-commands.png)

次の図は、ユーザーが Outlook でインライン応答を作成するときにアクティブ化される、アドイン コマンドが実装されていない 2 つの新規作成アドインが含まれたアドイン選択ウィンドウを示しています。

![作成されたアイテムに対してアクティブになるテンプレート メール アプリ。](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>新規作成モードで使用できるアドインの種類

新規作成アドインは [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)として実装されます。 メール作成または会議出席依頼の返信用のアドインをアクティブ化するために、アドインのマニフェストには [MessageComposeCommandSurface 拡張点要素](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface)が含まれます。 ユーザーが開催者である予定や会議の新規作成または編集を行うためのアドインをアクティブ化する場合、アドインには [AppointmentOrganizerCommandSurface 拡張点要素](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface)が含まれます。

> [!NOTE]
> アドイン コマンドがサポートされていないクライアントまたはサーバー用に開発されたアドインは、[OfficeApp](/javascript/api/manifest/officeapp) 要素に含まれる[ルール](/javascript/api/manifest/rule)要素の中の[アクティブ化ルール](activation-rules.md)を使用します。 アドインが特に古いクライアントやサーバーのために開発されている場合を除き、新規アドインはアドイン コマンドを使用すべきです。

## <a name="api-features-available-to-compose-add-ins"></a>新規作成アドインに使用できる API の機能

- [Outlook で新規作成フォームのアイテムに添付ファイルを追加および削除する](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する](get-set-or-add-recipients.md)
- [Outlook で予定またはメッセージを作成するときに件名を取得または設定する](get-or-set-the-subject.md)
- [Outlook で予定またはメッセージを作成するときに本文にデータを挿入する](insert-data-in-the-body.md)
- [Outlook で予定を作成するときに場所を取得または設定する](get-or-set-the-location-of-an-appointment.md)
- [Outlook で予定を作成するときに時刻を取得または設定する](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a>関連項目

- [Office の Outlook アドインの概要](../quickstarts/outlook-quickstart.md)

Outlook アドイン コードで Identity API セット 1.3 を要求するには、`isSetSupported('IdentityAPI', '1.3')` を呼び出してサポートされているかどうかを確認します。 Outlook アドインのマニフェストでの宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)」を参照してください。
export class OfficeUtils {
  static getSelectedSlideIndexAsync(): Promise<number> {
    return new Promise<number>((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        try {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(console.error(asyncResult.error.message));
          } else {
            resolve((asyncResult.value as any).slides[0].index);
          }
        } catch (error) {
          reject(console.log(error));
        }
      });
    });
  }
}

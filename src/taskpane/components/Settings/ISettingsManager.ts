export interface ISettingsManager {
  loadAsync(): Promise<IUnsplashAddinSettings>;
  // eslint-disable-next-line no-unused-vars
  save(settings: IUnsplashAddinSettings): void;
}

export class IUnsplashAddinSettings {
  insertUnsplashLogo: boolean = true;
  insertUnsplashAuthor: boolean = true;
  teachingBubbleNeverShown: boolean = true;
}

export class OfficeSettingsManager implements ISettingsManager {
  save(settings: IUnsplashAddinSettings): void {
    window.localStorage.clear();
    window.localStorage.setItem("settings-1.0", JSON.stringify(settings));
  }

  loadAsync(): Promise<IUnsplashAddinSettings> {
    return new Promise<IUnsplashAddinSettings>((resolve) => {
      const settings: IUnsplashAddinSettings = window.localStorage.getItem("settings-1.0")
        ? JSON.parse(window.localStorage.getItem("settings-1.0")!)
        : new IUnsplashAddinSettings();

      resolve(settings);
    });
  }
}

import * as React from "react";
import { DefaultButton, Icon, Panel, PanelType, SearchBox, Stack, TeachingBubble, Toggle } from "@fluentui/react";
import Progress from "../Progress";
import ImagesMasonry from "../ImagesMasonry/ImagesMasonry";
import { Unsplash } from "../ImagesMasonry/UnsplashDTOs";
import "./App.scss";
import { ISettingsManager, IUnsplashAddinSettings, OfficeSettingsManager } from "../Settings/ISettingsManager";

// import Header from "../Header";
// import HeroList, { HeroListItem } from "../HeroList";

/* global Office */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  // listItems: HeroListItem[];
  searchBoxText: string;
  teachingBubbleVisible: boolean;
  selectedImageCount: number;
  selectedImages: Unsplash.Image[];
  masonrySearchTerm: string;
  masonryHeight: string;
  settingsPanelVisible: boolean;
  settings: IUnsplashAddinSettings;
}

export default class App extends React.Component<AppProps, AppState> {
  pageSize = 30;
  unsplashClientId: string = process.env.REACT_APP_UNSPLASH_API_KEY as string;
  settingsManager: ISettingsManager;
  headerElement = React.createRef<HTMLDivElement>();
  resizeObserver?: ResizeObserver;

  constructor(props: AppProps) {
    super(props);

    this.state = {
      // listItems: [],
      selectedImageCount: 0,
      selectedImages: [],
      searchBoxText: "",
      masonrySearchTerm: "",
      teachingBubbleVisible: false,
      settingsPanelVisible: false,
      settings: new IUnsplashAddinSettings(),
      masonryHeight: "100hv",
    };

    this.settingsManager = new OfficeSettingsManager();

    this.settingsManager.loadAsync().then((settings) => {
      console.debug("✅ Settings loaded", settings);
      this.setState({ settings: settings });

      if (settings.teachingBubbleNeverShown) {
        console.debug("Showing teaching bubble once", settings);
        this.setState({ teachingBubbleVisible: true });

        settings.teachingBubbleNeverShown = false;
        this.settingsManager.save(settings);
      }
    });
  }

  componentDidMount() {
    console.debug("componentDidMount()");

    this.resizeObserver = new ResizeObserver((entries) => {
      console.debug("Resize Observer", entries);
      this.setState({ masonryHeight: `calc(100vh - ${entries[0].contentRect.height}px - 10px)` });
    });
    this.resizeObserver.observe(this.headerElement.current!);
  }

  componentWillUnmount() {
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
    }
  }

  toggleInsertLogoChanged = (_ev: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => {
    console.debug("Toggle Insert Unsplash Logo got clicked");

    let s = this.state.settings;
    s.insertUnsplashLogo = checked ?? true;
    this.setState({ settings: s });

    this.settingsManager.save(s);
  };

  toggleInsertAuthorChanged = (_ev: React.MouseEvent<HTMLElement, MouseEvent>, checked?: boolean) => {
    console.debug("Toggle Insert Unsplash Author got clicked");

    let s = this.state.settings;
    s.insertUnsplashAuthor = checked ?? true;
    this.setState({ settings: s });

    this.settingsManager.save(s);
  };

  btnResetSettings = () => {
    console.debug("btnResetSettings got clicked");

    const s = new IUnsplashAddinSettings();
    this.setState({ settings: s });
    this.settingsManager.save(s);
  };

  btnSearchClick = () => {
    console.debug("btnSearch got clicked");
    this.setState({ masonrySearchTerm: this.state.searchBoxText });
  };

  btnSettingsClick = () => {
    console.debug("btnSettingsClick got clicked");
    this.setState({ settingsPanelVisible: true });
  };

  handleSelectedImagesChanged = (images: Unsplash.Image[]) => {
    this.setState({ selectedImages: images });
    this.setState({ selectedImageCount: images.length });
  };

  btnInsertClick = () => {
    console.debug("btnInsert got clicked");
    const selectedImages = this.state.selectedImages;

    // Insert all selected images
    Promise.all(
      selectedImages.map((v) => {
        this.insertImageAsync(v).then(() => {
          const si = this.state.selectedImages.filter((p) => p != v);
          this.setState({ selectedImages: si });
        });
      })
    ).then(() => {
      if (this.state.settings.insertUnsplashLogo) {
        // Insert Unsplash logo
        this.getBase64ImageAsync("/assets/icon-128.png").then((unsplashLogo) => {
          console.debug("Inserting unsplash logo");
          Office.context.document.setSelectedDataAsync(
            unsplashLogo,
            {
              coercionType: Office.CoercionType.Image,
              imageLeft: 0,
              imageTop: 0,
              imageWidth: 100,
            },
            (asyncResult) => {
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("❎ Error inserting unsplash logo", asyncResult.error.message);
              } else {
                console.debug("✅ Successfully inserted unsplash logo");
              }
            }
          );
        });
      }

      console.debug("Resetting selected images to 0");
      this.setState({ selectedImageCount: 0 });
      this.setState({ selectedImages: [] });
    });
  };

  getBase64ImageAsync = (url: string) => {
    return new Promise<string>((resolve, reject) => {
      console.debug(`Downloading ${url} as base64`);
      var xhr = new XMLHttpRequest();
      xhr.onload = function () {
        var reader = new FileReader();
        reader.onloadend = () => {
          console.debug(`✅ Downloaded ${url} as base64 successfully`);
          const regex = /data:image\/[a-z0-9]*;base64,/i;
          const base64Image = (reader.result as string).replace(regex, "");
          resolve(base64Image);
        };
        reader.onerror = (err) => {
          console.error(`❎ Error downloading ${url} as base64`);
          reject(err);
        };
        reader.readAsDataURL(xhr.response);
      };
      xhr.open("GET", url);
      xhr.responseType = "blob";
      xhr.send();
    });
  };

  insertIntoPptAsync = (image: string, options: Office.SetSelectedDataOptions) => {
    return new Promise<void>((resolve, reject) => {
      console.debug(`Inserting image into document`);
      Office.context.document.setSelectedDataAsync(image, options, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`❎ Error inserting image into Office document:`, asyncResult.error.message);
          reject(asyncResult.error.message);
        } else {
          console.debug(`✅ Inserted image into Office document successfully`);
          resolve();
        }
      });
    });
  };

  insertImageAsync = (image: Unsplash.Image) => {
    return new Promise<void>((resolve, reject) => {
      console.debug("Downloading & inserting image, credit and logo");

      this.getBase64ImageAsync(image.urls.full)
        .then((base64Image: string) => {
          // links.download_location
          this.insertIntoPptAsync(base64Image, {
            coercionType: Office.CoercionType.Image,
            // imageLeft: 50,
            // imageTop: 50
            // imageWidth: 400
          })
            .catch((reason) => reject(reason))
            .then(() => {
              console.debug(`Tracking download on ${image.links.download_location}`);
              // trigger download link
              fetch(image.links.download_location + `&client_id=${this.unsplashClientId}`).catch((reason) =>
                console.error("❎ Error tracking download of the image", reason)
              );

              if (this.state.settings.insertUnsplashAuthor) {
                // Insert Unsplash author reference
                console.debug(`Inserting Unsplash author reference`);
                // const credit = `Photo by <a href="https://unsplash.com/@${image.user.username}?utm_source=your_app_name&utm_medium=referral">${image.user.name}</a> on <a href="https://unsplash.com/?utm_source=your_app_name&utm_medium=referral">Unsplash</a>`;
                const credit = `Photo by ${image.user.name} (https://unsplash.com/@${image.user.username}) on Unsplash (https://unsplash.com/)`;

                this.insertIntoPptAsync(credit, {
                  coercionType: Office.CoercionType.Text,
                })
                  .catch((reason) => reject(reason))
                  .then(() => {
                    console.debug("✅ Successfully Downloaded & inserted images (and maybe logo & credits)");
                    resolve();
                  });
              } else {
                console.debug("✅ Successfully Downloaded & inserted images (and maybe logo & credits)");
                resolve();
              }
            });
        })
        .catch((reason) => reject(reason));
    });
  };

  // enrichThumbSizes(results: any[]): any[] {
  //   return results.map(r => {
  //     r.thumbWidth = 200;
  //     r.thumbHeight = r.height / (r.width / 200);
  //     return r;
  //   });
  // }

  render() {
    console.debug("render()", this.state);
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("../../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    let selectedImagesHtml = "";
    if (this.state.selectedImageCount > 0) {
      selectedImagesHtml = `(${this.state.selectedImageCount})`;
    }

    return (
      <div id="container2">
        <Panel
          isOpen={this.state.settingsPanelVisible}
          onDismiss={() => this.setState({ settingsPanelVisible: false })}
          // headerText="Info & Settings"
          type={PanelType.smallFixedFar}
          isLightDismiss={true}
        >
          <h3>About</h3>
          <h4>Office Addin for Unsplash</h4>
          <p>Version: LOCAL-DEV</p>
          <p>© 2022 by knom</p>
          <p>
            <a href="https://github.com/knom/UnsplashOfficeJSAddin/issues" target="_blank" rel="noreferrer">
              Report a bug <Icon iconName="OpenInNewTab" />
            </a>
          </p>
          <h3>Settings</h3>
          <Toggle
            label="Insert Unsplash Logo"
            onText="On"
            offText="Off"
            checked={this.state.settings.insertUnsplashLogo}
            onChange={this.toggleInsertLogoChanged}
          />
          <Toggle
            label="Insert Unsplash Author References"
            onText="On"
            offText="Off"
            checked={this.state.settings.insertUnsplashAuthor}
            onChange={this.toggleInsertAuthorChanged}
          />
          <DefaultButton onClick={this.btnResetSettings}>Reset Settings</DefaultButton>
          <h3>Unsplash Usage Policy</h3>
          <a href="https://unsplash.com/license" target="_blank" rel="noreferrer">
            View Photo License <Icon iconName="OpenInNewTab" />
          </a>
        </Panel>

        <div id="header" ref={this.headerElement}>
          <Stack horizontal wrap tokens={{ childrenGap: 10, padding: 10 }}>
            <Stack.Item>
              <img src="assets/icon-32.png" alt="Unsplash logo" />
            </Stack.Item>
            <Stack.Item>
              <SearchBox
                id="searchBox"
                styles={{ root: { width: "180px" } }}
                showIcon={true}
                value={this.state.searchBoxText ?? ""}
                placeholder="Search for photos"
                onChange={(_ev, newValue) => this.setState({ searchBoxText: newValue === undefined ? "" : newValue })}
                onSearch={() => this.btnSearchClick()}
                autoFocus
              />
              {this.state.teachingBubbleVisible && (
                <TeachingBubble
                  target={`#searchBox`}
                  onDismiss={() => this.setState({ teachingBubbleVisible: false })}
                  headline="Type your favorite topic to get started..."
                  primaryButtonProps={{
                    text: "Get started",
                    onClick: () => this.setState({ teachingBubbleVisible: false }),
                  }}
                >
                  <p>
                    Then hit <b>ENTER</b> - and browse through the beautiful photos from <b>Unsplash</b>!
                  </p>
                  <p>
                    Once you&apos;re done browsing around, select your images and press <b>Insert</b> to add them to
                    your document.
                  </p>
                </TeachingBubble>
              )}
            </Stack.Item>
            {/* <Stack.Item>
              <PrimaryButton
                className="ms-Button ms-Button--primary"
                iconProps={{ iconName: "search" }}
                onClick={this.btnSearchClick}
                styles={{ root: { minWidth: 20, paddingRight: 4, paddingLeft: 4 } }}
              />
            </Stack.Item> */}
            <Stack.Item>
              <DefaultButton
                className="ms-Button"
                disabled={this.state.selectedImageCount == 0}
                onClick={this.btnInsertClick}
              >
                <span className="ms-Button-label">Insert {selectedImagesHtml}</span>
              </DefaultButton>
            </Stack.Item>
            <Stack.Item>
              <DefaultButton
                className="ms-Button"
                iconProps={{ iconName: "settings" }}
                onClick={this.btnSettingsClick}
                styles={{ root: { minWidth: 20, paddingRight: 4, paddingLeft: 4 } }}
              />
            </Stack.Item>
          </Stack>
        </div>
        <div style={{ height: this.state.masonryHeight }}>
          <ImagesMasonry
            searchTerm={this.state.masonrySearchTerm}
            onSelectedImagesChanged={this.handleSelectedImagesChanged}
            // showSelectedSpinner={this.state.showSelectedSpinner}
            selectedImages={this.state.selectedImages}
          />
        </div>
      </div>
    );
  }
}

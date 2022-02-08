import * as React from "react";
import { DefaultButton, PrimaryButton, SearchBox, Stack } from "@fluentui/react";
import Progress from "../Progress";
import ImagesMasonry from "../ImagesMasonry/ImagesMasonry";
import { Unsplash } from "../ImagesMasonry/UnsplashDTOs";
import "./App.scss";

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
  selectedImageCount: number;
  selectedImages: Unsplash.Image[];
  masonrySearchTerm: string;
}

export default class App extends React.Component<AppProps, AppState> {
  pageSize = 30;
  unsplashClientId: string = process.env.REACT_APP_UNSPLASH_API_KEY as string;

  constructor(props: AppProps) {
    super(props);

    this.state = {
      // listItems: [],
      selectedImageCount: 0,
      selectedImages: [],
      searchBoxText: "",
      masonrySearchTerm: "",
    };
  }

  componentDidMount() {
    console.debug("componentDidMount()");
  }

  btnSearchClick = () => {
    console.debug("btnSearch got clicked");
    this.setState({ masonrySearchTerm: this.state.searchBoxText });
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

          Office.addin.hide();
        });
      })
    ).then(() => {
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

              // const credit = `Photo by <a href="https://unsplash.com/@${image.user.username}?utm_source=your_app_name&utm_medium=referral">${image.user.name}</a> on <a href="https://unsplash.com/?utm_source=your_app_name&utm_medium=referral">Unsplash</a>`;
              const credit = `Photo by ${image.user.name} (https://unsplash.com/@${image.user.username}) on Unsplash (https://unsplash.com/)`;

              this.insertIntoPptAsync(credit, {
                coercionType: Office.CoercionType.Text,
              })
                .catch((reason) => reject(reason))
                .then(() => {
                  console.debug("✅ Successfully Downloaded & inserted image, credit and logo");
                  resolve();
                });
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
        <div id="header">
          <Stack horizontal wrap tokens={{ childrenGap: 10, padding: 10 }}>
            <Stack.Item>
              <img src="assets/icon-32.png" alt="Unsplash logo" />
            </Stack.Item>
            <Stack.Item>
              <SearchBox
                styles={{ root: { width: "180px" } }}
                value={this.state.searchBoxText ?? ""}
                placeholder="Search for photos"
                onChange={(_ev, newValue) => this.setState({ searchBoxText: newValue === undefined ? "" : newValue })}
                onSearch={() => this.btnSearchClick()}
              />
            </Stack.Item>
            <Stack.Item>
              <PrimaryButton className="ms-Button ms-Button--primary" onClick={this.btnSearchClick}>
                <span>
                  <i className="ms-Icon ms-Icon--Search searchIcon"></i>
                </span>
                &nbsp;
                <span className="ms-Button-label">Search</span>
              </PrimaryButton>
            </Stack.Item>
            <Stack.Item>
              <DefaultButton
                className="ms-Button"
                disabled={this.state.selectedImageCount == 0}
                onClick={this.btnInsertClick}
              >
                <span className="ms-Button-label">Insert {selectedImagesHtml}</span>
              </DefaultButton>
            </Stack.Item>
          </Stack>
        </div>
        <ImagesMasonry
          searchTerm={this.state.masonrySearchTerm}
          onSelectedImagesChanged={this.handleSelectedImagesChanged}
          // showSelectedSpinner={this.state.showSelectedSpinner}
          selectedImages={this.state.selectedImages}
        />
      </div>
    );
  }
}

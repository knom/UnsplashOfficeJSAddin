import * as React from "react";
import { DefaultButton, Icon, PrimaryButton, SearchBox, Stack } from "@fluentui/react";
import Header from "../Header";
import HeroList, { HeroListItem } from "../HeroList";
import Progress from "../Progress";
import ImagesMasonry from "../ImagesMasonry/ImagesMasonry";
import './App.scss';

/* global console, Office, require */

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

  constructor(props, context) {
    super(props, context);
    this.state = {
      // listItems: [],
      selectedImageCount: 0,
      selectedImages: [],
      searchBoxText: "",
      masonrySearchTerm: "",
    };
  }

  componentDidMount() {
    console.log("componentDidMount");

    const _this = this;
    // this.setState({
    //   listItems: [
    //     {
    //       icon: "Ribbon",
    //       primaryText: "Achieve more with Office integration",
    //     },
    //     {
    //       icon: "Unlock",
    //       primaryText: "Unlock features and functionality",
    //     },
    //     {
    //       icon: "Design",
    //       primaryText: "Create and visualize like a pro",
    //     },
    //   ],
    // });
  }

  btnSearchClick = () => {
    this.setState({ masonrySearchTerm: this.state.searchBoxText });
  }

  handleSelectedImagesChanged = (images: Unsplash.Image[]) => {
    this.setState({ selectedImages: images })
    this.setState({ selectedImageCount: images.length })
  };

  btnInsertClick = () => {
    const selectedImages = this.state.selectedImages;

    // Insert all selected images
    selectedImages.forEach((v) => {
      this.insertImageAsync(v).then(() => {
        const si = this.state.selectedImages.filter(p => p != v);
        this.setState({ selectedImages: si })
      });
    });

    // Insert Unsplash logo
    const unsplashLogo = "iVBORw0KGgoAAAANSUhEUgAAAf0AAAH9CAYAAAAQzKWIAAAACXBIWXMAAC4jAAAuIwF4pT92AAAFIGlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNS42LWMxNDUgNzkuMTYzNDk5LCAyMDE4LzA4LzEzLTE2OjQwOjIyICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyIgeG1sbnM6cGhvdG9zaG9wPSJodHRwOi8vbnMuYWRvYmUuY29tL3Bob3Rvc2hvcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgMjAxOSAoTWFjaW50b3NoKSIgeG1wOkNyZWF0ZURhdGU9IjIwMTgtMTAtMDNUMTI6Mzk6MDYtMDQ6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDE4LTEyLTE4VDE2OjAwOjUxLTA1OjAwIiB4bXA6TWV0YWRhdGFEYXRlPSIyMDE4LTEyLTE4VDE2OjAwOjUxLTA1OjAwIiBkYzpmb3JtYXQ9ImltYWdlL3BuZyIgcGhvdG9zaG9wOkNvbG9yTW9kZT0iMyIgcGhvdG9zaG9wOklDQ1Byb2ZpbGU9InNSR0IgSUVDNjE5NjYtMi4xIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOmExMDVjMjAwLWViYjctNDRlMy05YjA5LTExZjE4YjYxOTkyMCIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDphMTA1YzIwMC1lYmI3LTQ0ZTMtOWIwOS0xMWYxOGI2MTk5MjAiIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDphMTA1YzIwMC1lYmI3LTQ0ZTMtOWIwOS0xMWYxOGI2MTk5MjAiPiA8eG1wTU06SGlzdG9yeT4gPHJkZjpTZXE+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJjcmVhdGVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOmExMDVjMjAwLWViYjctNDRlMy05YjA5LTExZjE4YjYxOTkyMCIgc3RFdnQ6d2hlbj0iMjAxOC0xMC0wM1QxMjozOTowNi0wNDowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIENDIDIwMTkgKE1hY2ludG9zaCkiLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+S8e/QAAACDVJREFUeJzt18FxBDEMBDHR+efMi8LLRwMRzEOlLs7uPrgyMx4gKbs71xvo+rseAAB8Q/QBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBidvd6AwDwAZc+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AESIPgBEiD4ARIg+AETMe2+vR9C1u3O9Ab40M/5czrj0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIEL0ASBC9AEgQvQBIGLee3s9AgD4fy59AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiBB9AIgQfQCIEH0AiPgBtIET71uvWN4AAAAASUVORK5CYII=";
    Office.context.document.setSelectedDataAsync(
      unsplashLogo,
      {
        coercionType: Office.CoercionType.Image,
        imageLeft: 0,
        imageTop: 0
        // imageWidth: 400
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
        }
      }
    );
    this.setState({ selectedImageCount: 0 });
    this.setState({ selectedImages: [] });
  }

  getBase64ImageAsync = (url: string) => {
    return new Promise((resolve, reject) => {
      var xhr = new XMLHttpRequest();
      xhr.onload = function () {
        var reader = new FileReader();
        reader.onloadend = function () {
          resolve(reader.result);
        };
        reader.onerror = (err) => reject(err);
        reader.readAsDataURL(xhr.response);
      };
      xhr.open("GET", url);
      xhr.responseType = "blob";
      xhr.send();
    });
  }

  insertIntoPptAsync = (image: string, options: Office.SetSelectedDataOptions) => {
    return new Promise<void>((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        image,
        options,
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            reject(asyncResult.error.message);
          }
          else {
            resolve();
          }
        });
    });
  }

  insertImageAsync = (image: Unsplash.Image) => {
    return new Promise<void>((resolve, reject) => {
      this.getBase64ImageAsync(image.urls.full).then((base64Image: string) => {
        const regex = /data:image\/[a-z0-9]*;base64,/i;
        base64Image = base64Image.replace(regex, "");
        // links.download_location
        this.insertIntoPptAsync(
          base64Image,
          {
            coercionType: Office.CoercionType.Image,
            // imageLeft: 50,
            // imageTop: 50
            // imageWidth: 400
          }).catch((reason) => reject(reason))
          .then(() => {
            // fetch(image.links.download);

            // const credit = `Photo by <a href="https://unsplash.com/@${image.user.username}?utm_source=your_app_name&utm_medium=referral">${image.user.name}</a> on <a href="https://unsplash.com/?utm_source=your_app_name&utm_medium=referral">Unsplash</a>`;
            const credit = `Photo by ${image.user.name} (https://unsplash.com/@${image.user.username}) on Unsplash (https://unsplash.com/)`;

            this.insertIntoPptAsync(credit,
              {
                coercionType: Office.CoercionType.Text
              }).catch((reason) => reject(reason))
              .then(() => resolve());
          });
      }).catch((reason) => reject(reason));
    });
  }

  // enrichThumbSizes(results: any[]): any[] {
  //   return results.map(r => {
  //     r.thumbWidth = 200;
  //     r.thumbHeight = r.height / (r.width / 200);
  //     return r;
  //   });
  // }

  render() {
    console.log("render", this.state);
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
      // <div className="ms-welcome">
      //   <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
      //   <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
      //     <p className="ms-font-l">
      //       Modify the source files, then click <b>Run</b>.
      //     </p>
      //     <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
      //       Run
      //     </DefaultButton>
      //   </HeroList>
      // </div>
      <div id="container2">
        <div id="header">
          <img src="assets/icon-32.png" alt="Unsplash logo" />
          <SearchBox 
            value={this.state.searchBoxText ?? ""}
            placeholder="Search for photos"
            onChange={(_ev, newValue) => this.setState({ searchBoxText: newValue })}
            onSearch={() => this.btnSearchClick()}
          />
          <PrimaryButton className="ms-Button ms-Button--primary" onClick={this.btnSearchClick}>
            <span><i className="ms-Icon ms-Icon--Search" style={{ color: 'white', fontWeight: 'bold' }}></i></span>&nbsp;
            <span className="ms-Button-label">Search</span>
          </PrimaryButton>
          <DefaultButton className="ms-Button" disabled={this.state.selectedImageCount == 0} onClick={this.btnInsertClick}>
            <span className="ms-Button-label">Insert {selectedImagesHtml}</span>
          </DefaultButton>
        </div>
        <ImagesMasonry searchTerm={this.state.masonrySearchTerm}
          onSelectedImagesChanged={this.handleSelectedImagesChanged}
          // showSelectedSpinner={this.state.showSelectedSpinner}
          selectedImages={this.state.selectedImages} />
      </div>
    );
  }
}

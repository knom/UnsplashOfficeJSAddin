import * as React from "react";
import { renderToString } from "react-dom/server";
import Masonry from "react-masonry-component";
import InfiniteScroll from "infinite-scroll";
import imagesLoaded from "imagesloaded";
import CheckMarkSvg from "./checkmark.svg";
import { Icon } from "@fluentui/react";
import { Unsplash } from "../ImagesMasonry/UnsplashDTOs";
import "./ImagesMasonry.scss";
import { Utils } from "../../../Utils";
import { appInsights } from "../../../AppInsights";
export interface ImagesMasonryProps {
  searchTerm: string;
  selectedImages: Unsplash.Image[];
  // eslint-disable-next-line no-unused-vars
  onSelectedImagesChanged(images: Unsplash.Image[]): void;
}

export interface ImagesMasonryState {
  searchResults: Unsplash.Image[];
  selectedSearchResults: Unsplash.Image[];
  selectedImageCount: number;
}

export default class ImagesMasonry extends React.Component<ImagesMasonryProps, ImagesMasonryState> {
  masonryParent: React.RefObject<any>;
  infScroll!: InfiniteScroll<Unsplash.Response>;
  unsplashClientId: string = process.env.REACT_APP_UNSPLASH_API_KEY as string;
  pageSize = 30;

  constructor(props: ImagesMasonryProps) {
    super(props);
    this.state = {
      searchResults: [],
      selectedSearchResults: [],
      selectedImageCount: 0,
    };

    this.masonryParent = React.createRef<any>();
  }

  componentDidMount() {
    const masonry = this.masonryParent.current.masonry;
    console.debug("componentDidMount()");

    const _this = this;

    let gridDiv = masonry.element as HTMLDivElement;
    gridDiv.addEventListener("click", this.searchResultImageClick);

    this.infScroll = new InfiniteScroll(gridDiv, {
      append: ".grid_item",
      outlayer: masonry,
      elementScroll: "#imageList",
      // eslint-disable-next-line no-unused-vars
      path: function (this: InfiniteScroll<Unsplash.Image[]>) {
        return `https://api.unsplash.com/search/photos?query=${_this.props.searchTerm}&per_page=${_this.pageSize}&client_id=${_this.unsplashClientId}&page=${this.pageIndex}`;
      },
      // load response as JSON
      responseBody: "json",
      status: ".page-load-status",
      history: false,
    });

    this.infScroll.on("request", (r) => {
      console.debug("InfiniteScroll Request", r);
    });
    this.infScroll.on("error", (err) => {
      console.error("InfiniteScroll Error", err);
    });

    this.infScroll.on("load", (response) => {
      console.debug("InfiniteScroll Load");

      let list = _this.state.searchResults;
      list = list.concat(response.results);
      _this.setState({ searchResults: list });

      // compile body data into HTML
      let itemsHTML = response.results.map((v, i) => renderToString(getItemHTML(v, i))).join("");

      var tmp = document.implementation.createHTMLDocument("");
      tmp.body.innerHTML = itemsHTML;
      let items = Array.from(tmp.body.children);

      // append item elements
      imagesLoaded(items, () => {
        gridDiv.append(...items);
        _this.masonryParent.current.performLayout();
        this.infScroll!.isLoading = false;
      });
    });

    function getItemHTML(item: Unsplash.Image, idx: number) {
      let i = _this.pageSize * (_this.infScroll!.pageIndex - 2) + idx;

      return (
        <div key={i} data-index={i} className="grid_item">
          <img src={item.urls.thumb} alt={item.description} />
          <CheckMarkSvg className="checkmarkIcon hidden" />
          {/* <Spinner size={SpinnerSize.medium} className="hidden" /> */}
          <div className="overlay ms-fontSize-12">
            <a href={item.links.html} title="Open Image in Browser" target="_blank" rel="noreferrer">
              {Utils.truncateString(item.description, 30, "...") || "Photo"} by {item.user.name}
              &nbsp;
              <Icon iconName="OpenInNewTab" />
            </a>
          </div>
        </div>
      );
    }
  }

  componentDidUpdate(prevProps: Readonly<ImagesMasonryProps>) {
    console.debug("componentDidUpdate()", prevProps);

    if (prevProps.searchTerm != this.props.searchTerm) {
      console.debug("New search term --> clear masonry selection & rendering");
      this.clearSelection();
      this.resetMasonry();

      appInsights.trackEvent({ name: "Search" });
    }

    if (!Utils.isEqual(prevProps.selectedImages, this.props.selectedImages) && this.props.selectedImages.length == 0) {
      console.debug("Different image selection --> reset masonry selection");
      this.setState({ selectedSearchResults: this.props.selectedImages });

      this.clearSelection();
    }
  }

  private clearSelection() {
    console.debug("Clearing selection...");
    const melement = this.masonryParent.current.masonry.element;

    var elements = [...melement.querySelectorAll(".selected-Image")].map((el) => el.closest(".grid_item"));

    elements.forEach((el) => {
      const img_item = el.querySelector("img");
      img_item.classList.remove("selected-Image");

      const checkmark_item = el.querySelector(".checkmarkIcon");
      checkmark_item.classList.add("hidden");
    });

    this.props.onSelectedImagesChanged([]);
  }

  private resetMasonry() {
    console.debug("Resetting masonry to empty and reloading first page...");
    this.setState({ searchResults: [] });
    this.setState({ selectedImageCount: 0 });
    this.setState({ selectedSearchResults: [] });

    this.masonryParent.current.masonry.items = [];

    let gridDiv = this.masonryParent.current.masonry.element as HTMLDivElement;
    gridDiv.textContent = "";

    window.scroll(0, 0);

    this.infScroll.canLoad = true;
    this.infScroll.pageIndex = 1;
    this.infScroll.loadCount = 0;
    this.infScroll.loadNextPage();
  }

  private searchResultImageClick = async (e: MouseEvent) => {
    console.debug("SearchResult Image clicked...");
    const masonry = this.masonryParent.current.masonry;

    let el = e.target as HTMLElement;

    const grid_item = el.closest(".grid_item");

    const img_item = grid_item?.querySelector("img");
    img_item?.classList.toggle("selected-Image");

    const checkmark_item = grid_item?.querySelector(".checkmarkIcon");
    checkmark_item?.classList.toggle("hidden");

    // Get all selected images in HTML
    var elements = [...masonry.element.querySelectorAll(".selected-Image")];
    console.log(`${elements.length} images selected`);

    this.setState({ selectedImageCount: elements.length });

    // Get their IDs
    var ids = elements.map((val) => Number.parseInt(val.parentElement.dataset.index));

    // Get selected images from downloaded REST/JSON
    var selectedImages = this.state.searchResults.filter((_v, idx) => ids.indexOf(idx) !== -1);

    this.setState({ selectedSearchResults: selectedImages });

    this.props.onSelectedImagesChanged(selectedImages);
  };

  render() {
    console.debug("render()", this.state);

    return (
      <div id="imageList">
        <Masonry options={{ fitWidth: true, gutter: 5, itemSelector: ".grid_item" }} ref={this.masonryParent}></Masonry>
      </div>
    );
  }
}

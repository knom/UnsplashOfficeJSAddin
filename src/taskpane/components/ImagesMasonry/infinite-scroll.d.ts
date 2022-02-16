/* eslint-disable no-dupe-class-members */
/* eslint-disable no-unused-vars */
declare module "infinite-scroll" {
  class InfiniteScroll<T> {
    constructor(masonryContainer: HTMLElement, conf: IConfig);

    on(event: "append", handler: (body: T, path: string, items: any, response: Response) => void): void;
    on(event: "last", handler: (body: T, path: string) => void): void;
    on(event: "error", handler: (error: string, path: string, response: Response) => void): void;
    on(event: "request", handler: (path: string, fetchPromise: Promise<any>) => void): void;
    on(event: "load", handler: (data: T, path: string, response: Response) => void): void;

    loadNextPage(): void;

    isLoading: boolean;
    canLoad: boolean;

    loadCount: number;
    pageIndex: number;
  }

  interface IConfig {
    append: string;
    outlayer: any;
    elementScroll?: string;
    path: () => string;
    responseBody: "json" | "xml";
    status: string;
    history: boolean;
  }
  export = InfiniteScroll;
}

declare module "infinite-scroll"
{
    class InfiniteScroll {
        constructor(masonryContainer: HTMLElement, conf: Config);

        on(event: "request", handler: (request: any) => void): void;
        on(event: "load", handler: (data: Unsplash.Response, url: string) => void): void;

        loadNextPage(): void;

        isLoading: boolean;
        canLoad: boolean;

        loadCount: number;
        pageIndex: number;
    }

    class Config {
        append: string;
        outlayer: any;
        path: () => string;
        responseBody: "json" | "xml";
        status: string;
        history: boolean;
    }
    export = InfiniteScroll;
}
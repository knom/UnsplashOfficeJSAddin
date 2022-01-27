declare module "infinite-scroll"
{
    class InfiniteScroll {
        constructor(masonryContainer: HTMLElement, conf: Config);

        on(event: "append", handler: (body: Unsplash.Response, path: string, items: any, response: Response) => void): void;
        on(event: "last", handler: (body: Unsplash.Response, path: string) => void): void;
        on(event: "error", handler: (error: string, path: string, response: Response) => void): void;
        on(event: "request", handler: (path: string, fetchPromise: Promise<any>) => void): void;
        on(event: "load", handler: (data: Unsplash.Response, path: string, response: Response) => void): void;

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
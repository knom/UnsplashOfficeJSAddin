namespace Unsplash{

    export interface Urls {
        raw: string;
        full: string;
        regular: string;
        small: string;
        thumb: string;
    }
    
    export interface Links {
        self: string;
        html: string;
        download: string;
        download_location: string;
    }


    export interface UserLinks {
        self: string;
        html: string;
        photos: string;
        likes: string;
        portfolio: string;
        following: string;
        followers: string;
    }

    export interface ProfileImage {
        small: string;
        medium: string;
        large: string;
    }

    export interface Social {
        instagram_username: string;
        portfolio_url: string;
        twitter_username: string;
        paypal_email?: any;
    }

    export interface User {
        id: string;
        updated_at: Date;
        username: string;
        name: string;
        first_name: string;
        last_name: string;
        twitter_username: string;
        portfolio_url: string;
        bio?: any;
        location: string;
        links: UserLinks;
        profile_image: ProfileImage;
        instagram_username: string;
        total_collections: number;
        total_likes: number;
        total_photos: number;
        accepted_tos: boolean;
        for_hire: boolean;
        social: Social;
    }

    export interface Type {
        slug: string;
        pretty_slug: string;
    }

    export interface Category {
        slug: string;
        pretty_slug: string;
    }

    export interface Subcategory {
        slug: string;
        pretty_slug: string;
    }

    export interface Ancestry {
        type: Type;
        category: Category;
        subcategory: Subcategory;
    }

    export interface CoverPhoto {
        id: string;
        created_at: Date;
        updated_at: Date;
        promoted_at: Date;
        width: number;
        height: number;
        color: string;
        blur_hash: string;
        description: string;
        alt_description: string;
        urls: Urls;
        links: Links;
        categories: any[];
        likes: number;
        liked_by_user: boolean;
        current_user_collections: any[];
        sponsorship?: any;
        topic_submissions: any;
        user: User;
    }

    export interface Source {
        ancestry: Ancestry;
        title: string;
        subtitle: string;
        description: string;
        meta_title: string;
        meta_description: string;
        cover_photo: CoverPhoto;
    }

    export interface Tag {
        type: string;
        title: string;
        source: Source;
    }

    export interface Image {
        id: string;
        created_at: Date;
        updated_at: Date;
        promoted_at: Date;
        width: number;
        height: number;
        color: string;
        blur_hash: string;
        description: string;
        alt_description: string;
        urls: Urls;
        links: Links;
        categories: any[];
        likes: number;
        liked_by_user: boolean;
        current_user_collections: any[];
        sponsorship?: any;
        topic_submissions: any;
        user: User;
        tags: Tag[];
    }

    export interface Response {
        total: number;
        total_pages: number;
        results: Image[];
    }

}
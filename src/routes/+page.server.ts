import { dev } from "$app/environment";
import type { Actions } from "@sveltejs/kit";

export const actions = {

    default: async ({request}) => {
        const data = await request.formData();

        const files = data.getAll('files') as Array<File>

        let path = '/var/www/html/files/';

        if(dev){
            path = 'C:/temp/';
        }
    }
} satisfies Actions;
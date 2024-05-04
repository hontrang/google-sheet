/**
 * @internal
 */
export interface Config {
    consumerID: string;
    consumerSecret: string;
}

/**
 * replace consumerID and consumerSecret with your own keys
 */
export default {
  consumerID: '',
  consumerSecret: ''
} as unknown as Config
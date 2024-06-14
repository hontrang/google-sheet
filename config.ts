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
  consumerID: process.env.CONSUMERID,
  consumerSecret: process.env.CONSUMERSECRET
} as unknown as Config
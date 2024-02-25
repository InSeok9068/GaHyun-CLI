export interface DatabaseType {}

export const Table = {} as const;
export type Table = (typeof Table)[keyof typeof Table];

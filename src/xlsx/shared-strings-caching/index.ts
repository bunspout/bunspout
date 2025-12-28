export type { SharedStringsCachingStrategy } from './strategy';
export { InMemoryStrategy } from './in-memory-strategy';
export { FileBasedStrategy } from './file-based-strategy';
export { CachingStrategyFactory, IN_MEMORY_ENTRY_OVERHEAD_BYTES, MAX_NUM_STRINGS_PER_TEMP_FILE } from './factory';
export { getMemoryLimitInKB } from './memory-limit';


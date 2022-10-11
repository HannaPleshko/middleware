import { base64Encode, base64Decode } from '@hcllabs/openclientkeepcomponent';

/**
 * Class returning instance of information about sync state for SyncFolderItems 
 */
export class SyncStateInfo {
    private static stateSeparator = "##";
    protected _lastIndex = 0;
    protected _timestamp: string | undefined;

    constructor(timestamp: string, lastIndex: number) {
        this._timestamp = timestamp;
        this._lastIndex = lastIndex;
    }

    static parse(syncState: string | undefined ): SyncStateInfo | undefined {
        if (syncState) {
            let tStamp;
            let lIndex = 0;
                const stateInfo = base64Decode(syncState);
                const separatorIndex = stateInfo.indexOf(SyncStateInfo.stateSeparator);
                if (separatorIndex >= 0) {
                    tStamp = stateInfo.substring(0, separatorIndex);
                    if (tStamp.length === 0)
                        tStamp = undefined;
                    const lIdx = stateInfo.substring(separatorIndex + SyncStateInfo.stateSeparator.length);
                    if (lIdx && lIdx.length > 0) {
                        lIndex = parseInt(lIdx);
                    }
                }
                if (tStamp) {
                    return new SyncStateInfo(tStamp, lIndex);
                }
        }
        return undefined;
    }

    /**
     * Returns the index of items returned already sync.
     */
    get lastIndex(): number {
        return this._lastIndex;
    }

    /**
     * Set the idex of the last item synchronized for this folder.
     */
    set lastIndex(lastIndex: number) {
        this._lastIndex = lastIndex; 
    }

    /**
     * Returns the timestamp of the sync.
     */
    get timestamp(): string | undefined {
        return this._timestamp;
    }

    /**
     * Set the timestamp of the last sync.
     */
    set timestamp(timestamp: string | undefined) {
        this._timestamp = timestamp; 
    }

    getBase64SyncState(): string | undefined {
        return this.timestamp ? base64Encode(this.timestamp + SyncStateInfo.stateSeparator + this.lastIndex) : undefined;
    }

}
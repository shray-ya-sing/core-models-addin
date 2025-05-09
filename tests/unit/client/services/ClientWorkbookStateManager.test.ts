import { ClientWorkbookStateManager } from "../../../../src/client/services/ClientWorkbookStateManager";

describe("ClientWorkbookStateManager", () => {
  let manager: ClientWorkbookStateManager;

  beforeEach(() => {
    manager = new ClientWorkbookStateManager();
  });

  test("sets and gets active sheet name", () => {
    manager.setActiveSheetName("Sheet1");
    expect(manager.getActiveSheetName()).toBe("Sheet1");
  });

  describe("invalidateCache", () => {
    test("invalidates full cache when no operation type", () => {
      // Set private fields
      (manager as any).cachedState = { dummy: true };
      (manager as any).lastCaptureTime = 123;
      const spy = jest.spyOn((manager as any).metadataCache, "invalidateAllChunks");

      manager.invalidateCache();

      expect((manager as any).cachedState).toBeNull();
      expect((manager as any).lastCaptureTime).toBe(0);
      expect(spy).toHaveBeenCalled();
    });

    test("does not invalidate for UI-only operations", () => {
      (manager as any).cachedState = { dummy: true };
      (manager as any).lastCaptureTime = 123;
      const spy = jest.spyOn((manager as any).metadataCache, "invalidateAllChunks");

      manager.invalidateCache("set_gridlines");

      expect((manager as any).cachedState).toEqual({ dummy: true });
      expect((manager as any).lastCaptureTime).toBe(123);
      expect(spy).not.toHaveBeenCalled();
    });

    test("invalidates for data-modifying operations", () => {
      (manager as any).cachedState = { dummy: true };
      (manager as any).lastCaptureTime = 123;
      const spy = jest.spyOn((manager as any).metadataCache, "invalidateAllChunks");

      manager.invalidateCache("rename_sheet");

      expect((manager as any).cachedState).toBeNull();
      expect((manager as any).lastCaptureTime).toBe(0);
      expect(spy).toHaveBeenCalled();
    });
  });

  describe("getCachedOrCaptureState", () => {
    test("captures fresh state when none cached", async () => {
      const captureSpy = jest
        .spyOn(manager as any, "captureWorkbookState")
        .mockResolvedValue({ state: true } as any);

      const state = await manager.getCachedOrCaptureState();
      expect(captureSpy).toHaveBeenCalledTimes(1);
      expect(state).toEqual({ state: true });
    });

    test("uses cache when not expired", async () => {
      const captureSpy = jest
        .spyOn(manager as any, "captureWorkbookState")
        .mockResolvedValue({ state: true } as any);

      const first = await manager.getCachedOrCaptureState();
      const second = await manager.getCachedOrCaptureState();
      expect(captureSpy).toHaveBeenCalledTimes(1);
      expect(second).toBe(first);
    });

    test("forces refresh when requested", async () => {
      const captureSpy = jest
        .spyOn(manager as any, "captureWorkbookState")
        .mockResolvedValue({ state: true } as any);

      await manager.getCachedOrCaptureState();
      await manager.getCachedOrCaptureState(true);
      expect(captureSpy).toHaveBeenCalledTimes(2);
    });
  });
});

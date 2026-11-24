export interface HistoryState<T> {
  past: T[];
  present: T;
  future: T[];
  maxDepth: number;
}

export function createHistory<T>(initial: T, maxDepth = 100): HistoryState<T> {
  return {
    past: [],
    present: initial,
    future: [],
    maxDepth,
  };
}

export function pushHistory<T>(state: HistoryState<T>, next: T): HistoryState<T> {
  if (Object.is(state.present, next)) {
    return state;
  }
  const past = [...state.past, state.present];
  if (past.length > state.maxDepth) {
    past.shift();
  }
  return {
    ...state,
    past,
    present: next,
    future: [],
  };
}

export function undoHistory<T>(state: HistoryState<T>): HistoryState<T> {
  if (state.past.length === 0) return state;
  const previous = state.past[state.past.length - 1];
  const past = state.past.slice(0, -1);
  return {
    ...state,
    past,
    present: previous,
    future: [state.present, ...state.future],
  };
}

export function redoHistory<T>(state: HistoryState<T>): HistoryState<T> {
  if (state.future.length === 0) return state;
  const [next, ...rest] = state.future;
  return {
    ...state,
    past: [...state.past, state.present],
    present: next,
    future: rest,
  };
}

export function canUndo<T>(state: HistoryState<T>): boolean {
  return state.past.length > 0;
}

export function canRedo<T>(state: HistoryState<T>): boolean {
  return state.future.length > 0;
}

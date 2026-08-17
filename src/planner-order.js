const compareOrderHints = (left, right) => {
  const leftHint = String(left.orderHint || '');
  const rightHint = String(right.orderHint || '');

  if (!leftHint && !rightHint) return 0;
  if (!leftHint) return 1;
  if (!rightHint) return -1;
  return leftHint < rightHint ? -1 : leftHint > rightHint ? 1 : 0;
};

export const orderPlannerBuckets = buckets => buckets
  .map((bucket, index) => ({ bucket, index }))
  .sort((left, right) => compareOrderHints(left.bucket, right.bucket) || left.index - right.index)
  .map(({ bucket }) => bucket);

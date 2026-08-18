import test from 'node:test';
import assert from 'node:assert/strict';
import { orderPlannerBuckets } from './planner-order.js';

test('orders Planner buckets from left to right by descending orderHint', () => {
  const buckets = [
    { id: 'ministros', orderHint: '100' },
    { id: 'julgamento', orderHint: '300' },
    { id: 'precedentes', orderHint: '200' },
  ];

  assert.deepEqual(orderPlannerBuckets(buckets).map(bucket => bucket.id), ['julgamento', 'precedentes', 'ministros']);
});

test('keeps buckets without orderHint stable and places them last', () => {
  const buckets = [
    { id: 'missing-a' },
    { id: 'ordered', orderHint: '100' },
    { id: 'missing-b', orderHint: '' },
  ];

  assert.deepEqual(orderPlannerBuckets(buckets).map(bucket => bucket.id), ['ordered', 'missing-a', 'missing-b']);
});

test('does not mutate the Graph response', () => {
  const buckets = [{ id: 'second', orderHint: '200' }, { id: 'first', orderHint: '100' }];

  orderPlannerBuckets(buckets);

  assert.deepEqual(buckets.map(bucket => bucket.id), ['second', 'first']);
});

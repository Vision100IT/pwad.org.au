import {OccasionGroupedById} from '../_gen-types';
import sanity from './_sanity';

export async function findAll(): Promise<OccasionGroupedById[]> {
  const query = ['*'].concat([
    '[_type=="occasion" && !defined(parent)]',
    'order(name asc)',
    `{
    _id,
    name,
    "values": *[_type=="occasion" && references(^._id)] | order(name asc) | { _id, name }
  }`
  ]);

  const result = (await sanity.fetch(query.join('|'))) as OccasionGroupedById[];

  const grouped = [];
  const emptyGroup = [];

  for (const item of result) {
    if (item.values.length === 0) {
      emptyGroup.push(item);
    } else {
      grouped.push(item);
    }
  }

  return [...grouped, {name: null, values: emptyGroup}];
}

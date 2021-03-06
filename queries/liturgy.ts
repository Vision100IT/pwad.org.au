import sanity from '../lib/sanity';
import {Liturgy} from './_types';

export async function getById(id: string): Promise<Liturgy> {
  return sanity.fetch(
    ['*']
      .concat([
        '[_type == "liturgy" && _id == $id][0]',
        `{
          _id,
          _type,
          title,
          files->{
            _id,
            _type,
            name,
            size,
            url
          },
          author->{
            _id,
            name,
            dates
          },
          content,
          keywords[]->{
            _id,
            name
          },
          categories[]->{
            _id,
            name
          }
        }`
      ])
      .join('|'),
    {id}
  );
}

type SearchInput = {
  keyword?: string;
  occasion?: string;
  textContains?: string;
};

export async function search({
  keyword,
  occasion,
  textContains
}: SearchInput): Promise<Liturgy[]> {
  const variables: Record<string, unknown> = {};
  let query = ['*'].concat(['[_type == "liturgy"]']);

  if (textContains) {
    variables.search = `${textContains}*`;
    query = query.concat([
      '[(title match $search || content match $search || keywords[]->name match $search)]'
    ]);
  }

  if (occasion) {
    variables.occasion = occasion;
    query = query.concat(['[occasion == $occasion]']);
  }

  if (keyword) {
    variables.keyword = keyword;
    query = query.concat(['[references($keyword)]']);
  }

  query = query.concat([
    '{_id, _type, title, content[0...1], keywords[]->{_id,name}}'
  ]);

  return sanity.fetch(query.join('|'), variables);
}

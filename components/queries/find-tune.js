import gql from 'graphql-tag';

export default gql`
  query findTune($title: String, $skip: Int, $limit: Int) {
    tuneMany(
      filter: {text_contains: $title}
      limit: $limit
      skip: $skip
      sort: TITLE_ASC
    ) {
      _id
      title
    }
  }
`;

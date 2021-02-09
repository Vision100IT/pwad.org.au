import {gql} from '@apollo/client';

export default gql`
  mutation UpdatePresentationOptions($input: PresentationOptionsInput!) {
    updatePresentationOptions(input: $input) {
      _id
      hasPaidAccount
      hasFreeAccount
      picture
      name {
        first
        last
      }
      periodEndDate
      stripeCustomerId
      shortlist {
        ... on Document {
          _id
          _type
        }
        ... on Hymn {
          title
          hymnNumber
        }
        ... on Liturgy {
          title
        }
        ... on Prayer {
          title
        }
        ... on Scripture {
          title
        }
      }
      presentationOptions {
        font
        background
        ratio
      }
    }
  }
`;

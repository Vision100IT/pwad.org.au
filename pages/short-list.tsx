import React from 'react';
import {NextPage} from 'next';
import {Text} from 'theme-ui';

import withApollo from '../lib/with-apollo-client';
import PageLayout from '../components/page-layout';
import ContentWrap from '../components/content-wrap';
import Link, {hymnLinkProps} from '../components/link';
import ShortListButton from '../components/shortlist-button';

import {useMeQuery} from '../components/queries';

const ShortList: NextPage = () => {
  const {data} = useMeQuery();

  return (
    <PageLayout>
      <Text as="h1" fontWeight="extraBold">
        Public Worship and Aids to Devotion Committee Website
      </Text>
      <ContentWrap>
        <Text as="h2">My Short List</Text>

        {data?.me && (
          <Text as="ul" appearance="prose">
            {data.me.shortlist.map(hymn => (
              <li key={hymn._id}>
                <ShortListButton hymn={hymn} />{' '}
                <Link {...hymnLinkProps(hymn)} />
              </li>
            ))}
          </Text>
        )}
      </ContentWrap>
    </PageLayout>
  );
};

export default withApollo(ShortList);
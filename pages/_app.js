import App, {Container} from 'next/app';
import Head from 'next/head';
import React from 'react';
import {ApolloProvider} from 'react-apollo';
import {createTheme, ThemeProvider} from 'mineral-ui/themes';
import Box from 'mineral-ui/Box';
import {flow} from 'lodash';
import withEmotion from '../lib/with-emotion';
import withApolloClient from '../lib/with-apollo-client';
import NavBar from '../components/nav-bar/nav-bar';
import Footer from '../components/footer/footer';

const blue = {
  10: '#CFE4FC',
  20: '#C2DDFC',
  30: '#B5D7FC',
  40: '#A6CFFC',
  50: '#8CC1FB',
  60: '#66ADFA',
  70: '#4379B4',
  80: '#2C5786',
  90: '#20456E',
  100: '#133153',
  inflection: 70
};

const theme = createTheme({
  colors: {
    theme: blue
  },
  overrides: {
    bp_narrow: '(min-width: 512px)',
    bp_medium: '(min-width: 768px)',
    bp_wide: '(min-width: 1024px)',

    color_theme_hover: blue[70],

    fontFamily: 'cabin'
  }
});

const enhance = flow(
  withEmotion,
  withApolloClient
);

class MyApp extends App {
  static async getInitialProps({Component, ctx}) {
    let pageProps = {};

    if (Component.getInitialProps) {
      pageProps = await Component.getInitialProps(ctx);
    }

    return {pageProps};
  }

  render() {
    const {Component, pageProps, apolloClient} = this.props;
    return (
      <Container>
        <Head>
          <title>Public Worship and Aids to Devotion - Home</title>
        </Head>
        <ApolloProvider client={apolloClient}>
          <ThemeProvider theme={theme}>
            <Box width={3 / 4} marginVertical={0} marginHorizontal="auto">
              <NavBar />
              <Component {...pageProps} />
              <Footer />
            </Box>
          </ThemeProvider>
        </ApolloProvider>
      </Container>
    );
  }
}

export default enhance(MyApp);
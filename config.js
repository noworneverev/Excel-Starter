const config = {
  gatsby: {
    pathPrefix: '/',
    siteUrl: 'https://excel-starter.netlify.app/', //
    gaTrackingId: 'UA-175514291-1',
    trailingSlash: false,
  },
  header: {
    logo: 'https://raw.githubusercontent.com/noworneverev/noworneverev.github.io/df0f5c444077b6848f16e9cd04fc1d66c0a03d7a/images/excel30-10.svg',     //https://graphql-engine-cdn.hasura.io/learn-hasura/assets/homepage/brand.svg //https://raw.githubusercontent.com/noworneverev/noworneverev.github.io/c1b09a0adee201e92a9418c2127e5a3fb3d6c80b/images/excel20-20.svg
    logoLink: '',    ///introduction'
    title: '',
      // "<a href='https://hasura.io/learn/'><img class='img-responsive' src='https://graphql-engine-cdn.hasura.io/learn-hasura/assets/homepage/learn-logo.svg' alt='Learn logo' /></a>",
    githubUrl: 'https://github.com/noworneverev/Excel-Starter',
    helpUrl: '',
    tweetText: '',
    social: `<li>
      <a href="https://www.facebook.com/AccodingTW" target="_blank" rel="noopener">
        <div class="twitterBtn">
          <img src='https://upload.wikimedia.org/wikipedia/commons/9/9b/Font_Awesome_5_brands_facebook-square.svg' alt={'Facebook'}/>
        </div>
      </a>
    </li>`,
    links: [{ text: '', link: '' }],
    search: {
      enabled: false,
      indexName: '',
      algoliaAppId: process.env.GATSBY_ALGOLIA_APP_ID,
      algoliaSearchKey: process.env.GATSBY_ALGOLIA_SEARCH_KEY,
      algoliaAdminKey: process.env.ALGOLIA_ADMIN_KEY,
    },
  },
  sidebar: {
    forcedNavOrder: [      
      '/intro-to-excel',
      '/formula',
      '/tricks',
      '/pivot-table',
      '/split-text',
      '/restore-file',
      '/custom-shortcut',
      '/macro',
      '/side-project',
    ],
    // collapsedNav: [
    //   '/chapter1',
    //   '/codeblock', // add trailing slash if enabled above
    // ],
    links: [{ text: 'Github', link: 'https://github.com/noworneverev/Excel-Starter' }, {text:'Facebook', link:'https://www.facebook.com/AccodingTW'}],
    frontline: false,
    ignoreIndex: true,    
    // title:
    //   "<a href='https://hasura.io/learn/'>graphql </a><div class='greenCircle'></div><a href='https://hasura.io/learn/graphql/react/introduction/'>react</a>",
  },
  siteMetadata: {
    title: 'Excel Tutorial for Accounting Firm Newbie',
    description: 'Excel Tutorial for Beginners',
    ogImage: null,
    docsLocation: 'https://github.com/noworneverev/Excel-Starter/tree/master/content', //https://github.com/hasura/gatsby-gitbook-boilerplate/tree/master/content
    favicon: 'https://i.imgur.com/IOMAeE3.jpg', // https://graphql-engine-cdn.hasura.io/img/hasura_icon_black.svg
  },
  pwa: {
    enabled: false, // disabling this will also remove the existing service worker.
    manifest: {
      name: 'Gatsby Gitbook Starter',
      short_name: 'GitbookStarter',
      start_url: '/',
      background_color: '#6b37bf',
      theme_color: '#6b37bf',
      display: 'standalone',
      crossOrigin: 'use-credentials',
      icons: [
        {
          src: 'src/pwa-512.png',
          sizes: `512x512`,
          type: `image/png`,
        },
      ],
    },
  },
};

module.exports = config;

module.exports = {
    entry: './build/app.js',
    output: {
        filename: './public/scripts/app.js'
    },
    module: {
        loaders: [
            { test: /\.html$/, loader: 'html' }
        ]
  }
};

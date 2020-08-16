import React from 'react';

const CommunityAuthor = ({ name, imageUrl, mailUrl, githubUrl, description }) => {
  return (
    <>
      {/* <h2 className="communitySection">關於作者</h2> */}
      <div className="authorSection">
        <div className="authorImg">
          <img src={imageUrl} alt={name} />
        </div>
        <div className="authorDetails">
          <div className="authorName">
            <strong>{name}</strong>
            {mailUrl ? (
              <a href={mailUrl} target="_blank" rel="noopener noreferrer">
                <img
                  src="https://image.flaticon.com/icons/svg/561/561127.svg"
                  alt="Mail Icon"
                  aria-label="Email"
                />
              </a>
            ) : null}
            {githubUrl ? (
              <a href={githubUrl} target="_blank" rel="noopener noreferrer">
                <img
                  src="https://storage.googleapis.com/graphql-engine-cdn.hasura.io/learn-hasura/assets/social-media/github-icon.svg"
                  alt="Github Icon"
                  aria-label="Github"
                />
              </a>
            ) : null}
          </div>
          <div className="authorDesc">{description}</div>
        </div>
      </div>
    </>
  );
};

export default CommunityAuthor;

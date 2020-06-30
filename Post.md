# Managing a SharePoint site using Microsoft Graph API and Python

I think one downside of all the emails and automated reporting we do nowadays is that people tend to start using their emails as document storage. I also means that when someone else needs to be brought up to speed on a certain operation either all previous emails need to forwarded or hope that they are saved on a NAS. Even worse sometimes they are stored on personal drives.

One way in which I have been trying to prevent those issues is to commit to keeping up the SharePoint site for my department up to date, which includes copies of all historical automated reports that are sent out to executives. This allows for there to be a central repository of these reports, allows not technical personnel to search for them easily, and the day to day management doesn't need to be handled by me but just someone with a bit of experience (It also allows to use Flow and PowerApps for some stuff, but that's outside the scope of this post). Now SharePoint (an the whole Office 365) is not my favorite but you gotta work with what you have got, and I will admit that their API is pretty easy to use.

## Microsoft Graph API
### Setup

A [quick overview of Microsoft Graph](https://docs.microsoft.com/en-us/graph/overview) can be found on MS site. Basically its the gateway we can use to access all sorts of information in Office 365.

In order for us to use the API we must first register our application on the [Azure Portal](https://docs.microsoft.com/en-us/graph/tutorials/python?tutorial-step=2). The microsoft documentation is pretty good at explaning the process. The main information we will need at the end is
* Client ID
* Client Secret

With just that information we can then generate an Oauth2 Authorization token and start getting data from office 365. For example we could use post man to make all the auth calls.

## Postman Example
This is how we would do it from there

## SharePoint Class
However since I'm constantly sending reports, I decided that automating postman or writing a script for each one would be tedious so we will create a Sharepoint Class that will handle the authorization, the token storage and updating and the most common calls to the specific sharepoint site.

We will accomplish this using the [requests-oauthlib](https://pypi.org/project/requests-oauthlib/) library.

First as the documentation tells us to do we install the package

``` pip install requests requests_oauthlib ```

Now we need to initialize our object, in order for this to work we are going to have to distinct initialization cases. One were we do not have a pre-existing tokena and well need to generate it. Another were we do have a token and we are going to point it towards our storage (either a redis instance, or a raw JSON storage for cases because sometimes I might want to run it on a ancient PC laying around my office).

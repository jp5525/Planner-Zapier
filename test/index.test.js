const zapier = require('zapier-platform-core');

// Use this to make test calls into your app:
const App = require('../index');
const appTester = zapier.createAppTester(App);
zapier.tools.env.inject();
const { 
    appId, 
    clientSecret, 
    tenant,
    group: Group,
    plan: Plan,
    bucket: Bucket,
    assignee: Assignee,
    searchTitle,
    searchDescription,
    searchBadTitle,
    task: Task
} = process.env
const authData = { appId, clientSecret, tenant};

describe("Zapier Planner Functions", () =>{

    test("Create task with all display names.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, Assignee, Plan , Bucket, Title:"Clean Room" },
        };

        const results = await appTester(
            App.resources.task.create.operation.perform,
            bundle
          );

        expect(results).toMatchObject({ title:"Clean Room" });


    })

    test("Create bucket with all display names.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, Plan, name: "Un Do" },
        };

        const results = await appTester(
            App.resources.bucket.create.operation.perform,
            bundle
          );

        expect(results).toMatchObject({name:"Un Do"});


    })

    test("Create plan with all display names.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, name: "Tester Plan" },
        };

        const results = await appTester(
            App.resources.plan.create.operation.perform,
            bundle
          );

        expect(results).toMatchObject({title:"Tester Plan"});


    })

    // Removed Because App Auth does NOT support replying to threads...
    // test("Create post with all display names.", async ()=>{
    //     const bundle = {
    //         authData,
    //         inputData: {
    //             Group: "Mine",
    //             Plan: "Tester O One",
    //             Task: "Tile Tub",
    //             "Content": "DONE!!1!"
    //         },
    //     };

    //     const results = await appTester(
    //         App.resources.comment.create.operation.perform,
    //         bundle
    //       );

    //     expect(results).toMatchObject({title:"Tester Plan"});


    // })

    test("Search for a task based on a title.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, Plan, Title: searchTitle },
        };

        const results = await appTester(
            App.resources.task.search.operation.perform,
            bundle
          );
        
        expect(results[0]).toMatchObject({
            title: searchTitle,
            details:{
                description: searchDescription
            }
        });

    })

    test("Search for a task based on a description.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, Plan, Description: searchDescription },
        };

        const results = await appTester(
            App.resources.task.search.operation.perform,
            bundle
          );

        
        expect(results[0]).toMatchObject({
            title: searchTitle,
            details:{
                description: searchDescription
            }
        });

    })

    test("Search for a task that does not exsist based on a description and title.", async ()=>{
        const bundle = {
            authData,
            inputData: { Group, Plan, Description: searchDescription, Title:searchBadTitle },
        };

        const results = await appTester(
            App.resources.task.search.operation.perform,
            bundle
          );

        expect(results).toEqual([]);

    })

    test("Search/Find task by description and then update the title description and % complete.", async()=>{
        const searchBundle = {
            authData,
            inputData: { Group, Plan, Description: searchDescription },
        };

        const results = await appTester(
            App.resources.task.search.operation.perform,
            searchBundle
          );
        
        const {
            '@odata.etag': Task_Etag,
            details:{
                '@odata.etag': Task_Details_Etag
            }
        } = results[0];

        const bundle = {
            authData,
            inputData: {
                Group,
                Plan,
                Bucket,
                Task,
                Title: "Tile Bathroom.",
                Description: "Not just the tub, but everything there is.",
                PercentComplete: 55,
                Task_Etag,
                Task_Details_Etag
            },
        };

        const updateResults = await appTester(
            App.creates.task.operation.perform,
            bundle
        );

        expect(updateResults).toMatchObject({
            status: 204,
            details:{
                status: 204
            }
        })
    }, 10*1000)
})


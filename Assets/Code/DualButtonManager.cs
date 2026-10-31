using System.Collections;
using UnityEngine;

public class DualButtonManager : MonoBehaviour
{
    public float delayTime;

    private bool buttonA;
    private bool buttonB;
    private bool activated;

    public GameObject[] deactivateObjects; // 문
    public GameObject[] activateObjects;   // 이펙트

    public void SetButtonState(int buttonNumber, bool pressed)
    {
        if (buttonNumber == 0)
            buttonA = pressed;
        else if (buttonNumber == 1)
            buttonB = pressed;

        if (buttonA && buttonB && !activated)
        {
            activated = true;
            StartCoroutine(ActivateObjects());
        }
    }

    private IEnumerator ActivateObjects()
    {
        yield return new WaitForSeconds(delayTime);

        foreach (GameObject obj in deactivateObjects)
            obj.SetActive(false);

        foreach (GameObject obj in activateObjects)
            obj.SetActive(true);
            
    }
}
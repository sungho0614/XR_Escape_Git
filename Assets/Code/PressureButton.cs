using System.Collections.Generic;
using UnityEngine;

public class PressureButton : MonoBehaviour
{
    public DualButtonManager manager;
    public int buttonNumber;

    private readonly HashSet<Collider> colliders = new HashSet<Collider>();

    private void OnTriggerEnter(Collider other)
    {
        colliders.Add(other);
        manager.SetButtonState(buttonNumber, true);
    }

    private void OnTriggerExit(Collider other)
    {
        colliders.Remove(other);

        if (colliders.Count == 0)
            manager.SetButtonState(buttonNumber, false);
    }
}